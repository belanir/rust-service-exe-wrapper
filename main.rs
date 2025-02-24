use clap::{Parser, Subcommand};
use std::ffi::OsString;
use std::process::{Child, Command};
use std::sync::mpsc::channel;
use std::time::Duration;
use windows_service::{
    service::{
        ServiceAccess, ServiceControlAccept, ServiceControlHandlerResult, ServiceErrorControl,
        ServiceInfo, ServiceStartType, ServiceStatus, ServiceState, ServiceType,
    },
    service_control_handler,
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
            // When no subcommand is provided, we assume the executable is running as a service.
            // The service control manager (SCM) will pass the launch arguments.
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
        // Pass the --name and --bat parameters so they are available when the service starts.
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

/// Entry point for the service. The SCM calls this function and passes the launch arguments.
fn service_main(arguments: Vec<OsString>) {
    // Parse the launch arguments provided by Windows SCM.
    let cli = Cli::parse_from(arguments);
    if let Err(e) = run_service(&cli) {
        eprintln!("Service error: {}", e);
    }
}

/// Runs the service logic. This function registers a control handler, spawns the batch file,
/// and waits for STOP, PAUSE, or CONTINUE commands. On STOP, it kills the batch process.
fn run_service(cli: &Cli) -> Result<(), Box<dyn std::error::Error>> {
    // Create a channel to receive control commands.
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
            // When STOP is received, kill the batch process.
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
            // For demonstration purposes, we simply log the pause event.
            eprintln!("Service paused. (Note: pausing a batch process is not directly supported.)");
            // You could implement additional logic here if you need to handle pausing.
        }
        "continue" => {
            // Log resume event.
            eprintln!("Service resumed.");
            // Implement resume logic as needed.
        }
        _ => {}
    }
    Ok(())
}
