use clap::{Parser, Subcommand, ValueHint};
use std::ffi::OsString;
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

/// Command-line arguments for the service manager.
/// The user must provide either `install` or `uninstall`.
#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Install the service.
    Install {
        #[arg(long, value_hint = ValueHint::Other, help = "Example: MyService")]
        name: String,

        #[arg(long, value_hint = ValueHint::FilePath, help = "Example: C:\\scripts\\run.bat")]
        bat: String,
    },
    /// Uninstall the service.
    Uninstall {
        #[arg(long, value_hint = ValueHint::Other, help = "Example: MyService")]
        name: String,
    },
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let cli = Cli::parse();

    match &cli.command {
        Commands::Install { name, bat } => {
            install_service(name, bat)?;
            println!("Service '{}' installed successfully.", name);
        }
        Commands::Uninstall { name } => {
            uninstall_service(name)?;
            println!("Service '{}' uninstalled successfully.", name);
        }
    }
    Ok(())
}

/// Installs the service with the given name and batch file path.
fn install_service(service_name: &str, bat_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let current_exe = std::env::current_exe()?;
    let service_info = ServiceInfo {
        name: OsString::from(service_name),
        display_name: OsString::from(service_name),
        service_type: ServiceType::OWN_PROCESS,
        start_type: ServiceStartType::Automatic,
        error_control: ServiceErrorControl::Normal,
        executable_path: current_exe.clone(),
        launch_arguments: vec![
            OsString::from("--name"),
            OsString::from(service_name),
            OsString::from("--bat"),
            OsString::from(bat_path),
        ],
        dependencies: vec![],
        account_name: None,
        account_password: None,
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

/// The service entry point. This is called by Windows when the service starts.
extern "system" fn service_main(argc: u32, argv: *mut *mut u16) {
    let args = raw_args_to_vec(argc, argv);
    let cli = Cli::parse_from(args);
    if let Err(e) = run_service(&cli) {
        eprintln!("Service error: {}", e);
    }
}

/// Converts raw Windows arguments into `Vec<OsString>`.
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

/// Runs the service logic: spawns the batch file and handles stop requests.
fn run_service(cli: &Cli) -> Result<(), Box<dyn std::error::Error>> {
    let (control_tx, control_rx) = channel();

    let event_handler = move |control_event| -> ServiceControlHandlerResult {
        match control_event {
            windows_service::service_control_handler::ServiceControl::Stop => {
                let _ = control_tx.send("stop");
                ServiceControlHandlerResult::NoError
            }
            _ => ServiceControlHandlerResult::NotImplemented,
        }
    };

    let status_handle = service_control_handler::register("MyRustService", event_handler)
        .expect("Failed to register service control handler");

    status_handle.set_service_status(ServiceStatus {
        service_type: ServiceType::OWN_PROCESS,
        current_state: ServiceState::Running,
        controls_accepted: ServiceControlAccept::STOP,
        checkpoint: 0,
        wait_hint: Duration::from_secs(5),
        process_id: None,
    })?;

    let mut child: Child = Command::new("cmd.exe")
        .args(&["/C", &cli.command.get_bat_path()])
        .spawn()?;

    let control_signal = control_rx.recv().unwrap();
    if control_signal == "stop" {
        let _ = child.kill();
        status_handle.set_service_status(ServiceStatus {
            service_type: ServiceType::OWN_PROCESS,
            current_state: ServiceState::Stopped,
            controls_accepted: ServiceControlAccept::empty(),
            checkpoint: 0,
            wait_hint: Duration::from_secs(5),
            process_id: None,
        })?;
    }
    Ok(())
}

trait GetBatPath {
    fn get_bat_path(&self) -> String;
}

impl GetBatPath for Commands {
    fn get_bat_path(&self) -> String {
        match self {
            Commands::Install { bat, .. } => bat.clone(),
            _ => String::new(),
        }
    }
}
