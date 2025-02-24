use clap::{Parser, Subcommand, ValueHint};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::mpsc::{channel, Receiver};
use std::thread;
use std::time::Duration;
use std::process::{Child, Command, Stdio};
use std::sync::{Arc, Mutex};
use tracing::{info, error, warn, debug};
use tracing_subscriber::fmt;
use tracing_subscriber::fmt::time::LocalTime;
use tracing_subscriber::EnvFilter;
use windows_service::{
    service::{
        ServiceAccess, ServiceControlAccept, ServiceErrorControl, ServiceInfo, ServiceStartType,
        ServiceStatus, ServiceState, ServiceType,
    },
    service_control_handler::{self, ServiceControlHandlerResult},
    service_dispatcher,
    service_manager::{ServiceManager, ServiceManagerAccess},
};

static CLI: once_cell::sync::Lazy<Arc<Mutex<Option<Cli>>>> = once_cell::sync::Lazy::new(|| Arc::new(Mutex::new(None)));

/// This function will store the `cli` object in the global state.
fn store_cli_object(cli: Cli) {
    let mut cli_lock = CLI.lock().unwrap();
    *cli_lock = Some(cli);
}

/// Initializes logging and writes logs in the same folder as the `.exe`
fn setup_logging() {
    let exe_path = std::env::current_exe().expect("Failed to get exe path");
    let log_path = exe_path.parent().unwrap_or_else(|| Path::new(".")).join("service.log");

    // Ensure log directory exists
    if let Some(parent) = log_path.parent() {
        fs::create_dir_all(parent).unwrap_or_else(|_| panic!("Failed to create log directory: {:?}", parent));
    }

    // Set up `tracing` subscriber
    tracing_subscriber::fmt()
        .with_writer(move || fs::OpenOptions::new().create(true).append(true).open(&log_path).unwrap())
        .with_timer(LocalTime::rfc_3339()) // Adds timestamps
        .with_env_filter(EnvFilter::new("debug")) // Default log level
        .init();
}

/// Command-line arguments for the service.
#[derive(Parser, Clone)]
#[command(author, version, about)]
struct Cli {
    #[arg(long, value_hint = ValueHint::Other, help = "Example: MyService")]
    name: String,

    #[arg(long, value_hint = ValueHint::FilePath, help = "Example: C:/scripts/run.bat")]
    bat: String,

    #[command(subcommand)]
    command: Option<Commands>,
}

#[derive(Subcommand, Clone)]
enum Commands {
    Install,
    Uninstall,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    setup_logging(); // ✅ Initialize logging

    let cli = Cli::parse();
    info!("Received CLI arguments: name='{}', bat='{}'", cli.name, cli.bat);

    match &cli.command {
        Some(Commands::Install) => {
            info!("Installing service '{}'", cli.name);
            install_service(&cli.name, &cli.bat)?;
            info!("Service '{}' installed successfully.", cli.name);
        }
        Some(Commands::Uninstall) => {
            info!("Uninstalling service '{}'", cli.name);
            uninstall_service(&cli.name)?;
            info!("Service '{}' uninstalled successfully.", cli.name);
        }
        None => {
            info!("Starting service '{}'", cli.name);
            store_cli_object(cli.clone());
            service_dispatcher::start(&cli.name, service_main)?;
        }
    }
    Ok(())
}

fn install_service(service_name: &str, bat_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let current_exe = std::env::current_exe()?;
    info!("Installing service '{}', using exe path: {:?}", service_name, current_exe);

    let service_info = ServiceInfo {
        name: service_name.into(),
        display_name: service_name.into(),
        service_type: ServiceType::OWN_PROCESS,
        start_type: ServiceStartType::OnDemand,
        error_control: ServiceErrorControl::Normal,
        executable_path: current_exe.clone(),
        launch_arguments: vec![
            "--name".into(),
            service_name.into(),
            "--bat".into(),
            bat_path.into(),
        ],
        dependencies: vec![],
        account_name: None,
        account_password: None,
    };

    let service_manager = ServiceManager::local_computer(
        None::<&str>,
        ServiceManagerAccess::CONNECT | ServiceManagerAccess::CREATE_SERVICE
    )?;
    let _service = service_manager.create_service(&service_info, ServiceAccess::empty())?;
    info!("Service '{}' installed successfully.", service_name);
    Ok(())
}

fn uninstall_service(service_name: &str) -> Result<(), Box<dyn std::error::Error>> {
    info!("Uninstalling service '{}'", service_name);
    let service_manager = ServiceManager::local_computer(None::<&str>, ServiceManagerAccess::CONNECT)?;
    let service = service_manager.open_service(service_name, ServiceAccess::DELETE)?;
    service.delete()?;
    info!("Service '{}' uninstalled successfully.", service_name);
    Ok(())
}

extern "system" fn service_main(argc: u32, argv: *mut *mut u16) {
    debug!("Starting service_main");

    let cli = {
        let cli_lock = CLI.lock().unwrap();
        cli_lock.as_ref().cloned()
    };

    if let Some(cli) = cli {
        info!("Executing service_main for '{}'", cli.name);
        if let Err(e) = run_service(&cli.name, &cli.bat) {
            error!("Service error: {}", e);
        }
    } else {
        error!("Failed to retrieve CLI arguments in service_main");
    }
}

fn raw_args_to_vec(argc: u32, argv: *mut *mut u16) -> Vec<String> {
    debug!("Starting raw args to vec");
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
            args.push(String::from_utf16_lossy(slice));
        }
    }
    args
}


/// Runs the service by starting the batch process in a child process. It then polls
/// for either a stop signal from the service control or the natural termination of the child.
fn run_service(service_name: &str, bat_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let (control_tx, control_rx) = channel();

    let event_handler = move |control_event| -> ServiceControlHandlerResult {
        if let windows_service::service::ServiceControl::Stop = control_event {
            let _ = control_tx.send("stop");
            return ServiceControlHandlerResult::NoError;
        }
        ServiceControlHandlerResult::NotImplemented
    };

    let status_handle = service_control_handler::register(service_name, event_handler)?;

    info!(
        "Service '{}' registered. Starting batch process...",
        service_name
    );
    status_handle.set_service_status(ServiceStatus {
        service_type: ServiceType::OWN_PROCESS,
        current_state: ServiceState::StartPending,
        controls_accepted: ServiceControlAccept::STOP,
        exit_code: Default::default(),
        checkpoint: 1,
        wait_hint: Duration::from_secs(10),
        process_id: Some(std::process::id()),
    })?;

    let mut child = Command::new("cmd.exe")
        .args(&["/K", bat_path])
        .stdout(Stdio::inherit())
        .stderr(Stdio::inherit())
        .spawn()?;

    info!("Batch file '{}' started successfully.", bat_path);

    status_handle.set_service_status(ServiceStatus {
        service_type: ServiceType::OWN_PROCESS,
        current_state: ServiceState::Running,
        controls_accepted: ServiceControlAccept::STOP,
        exit_code: Default::default(),
        checkpoint: 0,
        wait_hint: Duration::from_secs(5),
        process_id: Some(std::process::id()),
    })?;

    let stop_requested = wait_for_stop_signal(&control_rx, &mut child);

    if stop_requested {
        info!("Stop signal received; child process was terminated.");
    } else {
        info!("Child process finished naturally.");
    }

    info!("Service '{}' is stopping...", service_name);
    status_handle.set_service_status(ServiceStatus {
        service_type: ServiceType::OWN_PROCESS,
        current_state: ServiceState::Stopped,
        controls_accepted: ServiceControlAccept::empty(),
        exit_code: Default::default(),
        checkpoint: 0,
        wait_hint: Duration::from_secs(5),
        process_id: None,
    })?;
    info!("Service '{}' has stopped.", service_name);
    Ok(())
}

/// Polls for either a stop signal from the service control or for the child process
/// to finish naturally. No additional thread is spawned here.
fn wait_for_stop_signal(control_rx: &Receiver<&str>, child: &mut Child) -> bool {
    loop {
        // If a stop signal is received, kill the child process.
        if let Ok("stop") = control_rx.try_recv() {
            info!("Stop signal received from service control; terminating child process...");
            let _ = child.kill();
            let _ = child.wait();
            return true;
        }
        // If the child process has finished naturally, return immediately.
        if let Ok(Some(_)) = child.try_wait() {
            info!("Child process finished naturally.");
            return false;
        }
        info!("Service running...");
        thread::sleep(Duration::from_secs(1));
    }
}
