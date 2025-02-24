use clap::{Parser, Subcommand};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::process::{Child, Command};
use std::sync::mpsc::channel;
use std::time::Duration;
use windows_service::{
    service::{
        ServiceAccess, ServiceControlAccept, ServiceErrorControl, ServiceInfo, ServiceStartType,
        ServiceStatus, ServiceState, ServiceType,
    },
    service_control_handler::{self, ServiceControlHandlerResult},
    service_dispatcher,
    service_manager::{ServiceManager, ServiceManagerAccess},
};

/// Command-line arguments for the service executable.
/// Global options (--name and --bat) are used when running as a service.
/// Subcommands are used to install or uninstall the service.
#[derive(Parser)]
#[command(author, version, about)]
struct Cli {
    /// Name of the service.
    #[arg(long, default_value = "MyRustService")]
    name: String,

    /// Path to the batch file to run.
    #[arg(long, default_value = "C:\\path\\to\\default.bat")]
    bat: String,

    #[command(subcommand)]
    command: Option<Commands>,
}

#[derive(Subcommand)]
enum Commands {
    /// Install the service.
    Install {
        /// Service name.
        #[arg(long)]
        name: String,
        /// Path to the batch file.
        #[arg(long)]
        bat: String,
    },
    /// Uninstall the service.
    Uninstall {
        /// Service name.
        #[arg(long)]
        name: String,
    },
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let cli = Cli::parse();

    match &cli.command {
        Some(Commands::Install { name, bat }) => {
            install_service(name, bat)?;
            println!("Service '{}' installed successfully.", name);
        }
        Some(Commands::Uninstall { name }) => {
            uninstall_service(name)?;
            println!("Service '{}' uninstalled successfully.", name);
        }
        None => {
            // When no subcommand is provided, assume we're running as a service.
            // The service control manager (SCM) will pass raw launch arguments.
            service_dispatcher::start(&cli.name, service_main)?;
        }
    }
    Ok(())
}

/// Installs the service using the provided service name and batch file path.
/// The current executable is registered with Windows SCM along with launch arguments.
fn install_service(service_name: &str, bat_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let current_exe = std::env::current_exe()?;
    let service_info = ServiceInfo {
        name: OsString::from(service_name),
        display_name: OsString::from(service_name),
        service_type: ServiceType::OWN_PROCESS,
        start_type: ServiceStartType::Automatic,
        error_control: ServiceErrorControl::Normal,
        executable_path: current_exe.clone(),
        // Pass the --name and --bat parameters so they're available when the service starts.
        launch_arguments: vec![
            OsString::from("--name"),
            OsString::from(service_name),
            OsString::from("--bat"),
            OsString::from(bat_path),
        ],
        dependencies: vec![],
        account_name: None,     // Run as LocalSystem
        account_password: None, // Not used for LocalSystem
    };

    let manager_access = ServiceManagerAccess::CONNECT | ServiceManagerAccess::CREATE_SERVICE;
    let service_manager = ServiceManager::local_computer(None::<&str>, manager_access)?;
    let _service = service_manager.create_service(&service_info, ServiceAccess::empty())?;
    Ok(())
}

/// Uninstalls the service with the given name.
fn uninstall_service(service_name: &str) -> Result<(), Box<dyn std::error::Error>> {
    let manager_access = ServiceManagerAccess::CONNECT;
    let service_manager = ServiceManager::local_computer(None::<&str>, manager_access)?;
    let service = service_manager.open_service(service_name, ServiceAccess::DELETE)?;
    service.delete()?;
    Ok(())
}

/// The service entry point with the expected signature.
/// Converts raw Windows arguments into a Vec<OsString> for clap parsing.
extern "system" fn service_main(argc: u32, argv: *mut *mut u16) {
    let args = raw_args_to_vec(argc, argv);
    let cli = Cli::parse_from(args);
    if let Err(e) = run_service(&cli) {
        eprintln!("Service error: {}", e);
    }
}

/// Helper function to convert raw Windows arguments (wide strings) into a Vec<OsString>.
fn raw_args_to_vec(argc: u32, argv: *mut *mut u16) -> Vec<OsString> {
    let mut args = Vec::with_capacity(argc as usize);
    if argv.is_null() {
        return args;
    }
    for i in 0..argc {
        unsafe {
            let ptr = *argv.add(i as usize);
            if ptr.is_null() {
                continue;
            }
            // Find the length of the null-terminated wide string.
            let mut len = 0;
            while *ptr.add(len) != 0 {
                len += 1;
            }
            let slice = std::slice::from_raw_parts(ptr, len);
            args.push(OsString::from_wide(slice));
        }
    }
    args
}

/// Runs the service logic. Registers a control handler, spawns the batch file,
/// and waits for STOP, PAUSE, or CONTINUE commands. On STOP, the batch process is killed.
fn run_service(cli: &Cli) -> Result<(), Box<dyn std::error::Error>> {
    // Create a channel to receive control signals.
    let (control_tx, control_rx) = channel();

    let event_handler = move |control_event| -> ServiceControlHandlerResult {
        match control_event {
            windows_service::service_control_handler::ServiceControl::Stop => {
                let _ = control_tx.send("stop");
                ServiceControlHandlerResult::NoError
            }
            windows_service::service_control_handler::ServiceControl::Pause => {
                let _ = control_tx.send("pause");
                ServiceControlHandlerResult::NoError
            }
            windows_service::service_control_handler::ServiceControl::Continue => {
                let _ = control_tx.send("continue");
                ServiceControlHandlerResult::NoError
            }
            _ => ServiceControlHandlerResult::NotImplemented,
        }
    };

    // Register the service control handler.
    let status_handle = service_control_handler::register(&cli.name, event_handler)
        .expect("Failed to register service control handler");

    // Report that the service is running and accepts STOP and PAUSE/CONTINUE commands.
    status_handle.set_service_status(ServiceStatus {
        service_type: ServiceType::OWN_PROCESS,
        current_state: ServiceState::Running,
        controls_accepted: ServiceControlAccept::STOP | ServiceControlAccept::PAUSE_CONTINUE,
        checkpoint: 0,
        wait_hint: Duration::from_secs(5),
        process_id: None,
    })?;

    // Spawn the batch file via cmd.exe.
    let mut child: Child = Command::new("cmd.exe")
        .args(&["/C", &cli.bat])
        .spawn()?;

    // Wait for a control signal.
    let control_signal = control_rx.recv().unwrap();
    match control_signal {
        "stop" => {
            // Kill the batch process on STOP.
            let _ = child.kill();
            // Report that the service has stopped.
            status_handle.set_service_status(ServiceStatus {
                service_type: ServiceType::OWN_PROCESS,
                current_state: ServiceState::Stopped,
                controls_accepted: ServiceControlAccept::empty(),
                checkpoint: 0,
                wait_hint: Duration::from_secs(5),
                process_id: None,
            })?;
        }
        "pause" => {
            eprintln!("Service paused. (Pausing a batch process is not directly supported.)");
        }
        "continue" => {
            eprintln!("Service resumed.");
        }
        _ => {}
    }
    Ok(())
}
