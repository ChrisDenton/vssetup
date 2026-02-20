use std::ffi::OsString;
use std::io::Read;
use std::os::windows::ffi::OsStringExt;
use std::os::windows::fs::OpenOptionsExt;
use std::path::PathBuf;
use std::process::ExitCode;
use std::{env, process::Command};

use find_msvc_tools::find_windows_sdk;
use serde::Deserialize;
use vssetup::{HRESULT, SetupConfiguration, Variant, com, wide_str};

// channelId=VisualStudio.17.Release
// channelUri=https://aka.ms/vs/17/release/channel

#[derive(Deserialize)]
struct VsConfig {
    components: Vec<String>,
}

// TODO: properly handle errors.
fn instances() -> Vec<VisualStudio> {
    let result = Vec::new();

    let Ok(setup) = SetupConfiguration::new() else {
        return result;
    };
    let Ok(instances) = setup.EnumInstances() else {
        return result;
    };
    instances
        .map(|instance| {
            let product = instance.GetProduct().ok()??;
            let product_id = product.GetId().ok()?.to_string();
            let product_id = Product::from_str(&product_id)?;
            let install_path = OsString::from_wide(&instance.GetInstallationPath().ok()?).into();
            let display_name = instance.GetDisplayName(0x400).ok()?.to_string();
            //let id = instance.GetInstanceId().ok()?.to_string();
            let version = instance.GetInstallationVersion().ok()?.to_string();
            let major = version
                .split_once('.')
                .and_then(|v| v.0.parse::<u16>().ok())
                .unwrap_or_default();
            // Don't attempt to install for Visual Studio 2017 as that's untested.
            if major < 16 {
                return None;
            }

            let setup_path = instance
                .GetProperties()
                .ok()??
                .GetValue(wide_str!("setupEngineFilePath"))
                .ok()?;
            let setup_path: PathBuf = if let Variant::Bstr(setup_path) = setup_path {
                OsString::from_wide(&setup_path).into()
            } else {
                // Fallback to %ProgramFile(x86)%\Microsoft Visual Studio\installer\setup.exe.
                // If the "ProgramFiles(x86)" environment variable isn't set, then use the hardcoded value
                // (which will almost certainly be correct for user systems).
                let mut program_files = PathBuf::from(
                    env::var_os("ProgramFiles(x86)").unwrap_or(r"C:\Program Files (x86)".into()),
                );
                program_files.extend(["Microsoft Visual Studio", "installer", "setup.exe"]);
                program_files
            };

            /*let channel_id = instance
            .to_property_store()
            .ok()?
            .GetValue(wide_str!("channelId"))
            .ok()?
            .to_string();*/

            let mut msvc = None;
            let mut sdk = None;
            if let Ok(packages) = instance.GetPackages() {
                for package in packages.iter() {
                    if let Ok(id) = package.GetId() {
                        let component = Sdk::from_id(&id.to_string());
                        if component.is_some() && component > sdk {
                            sdk = component;
                        } else if let Some(component) = Msvc::from_id(&id.to_string())
                            && component == Msvc::from_rust_arch(std::env::consts::ARCH)
                        {
                            msvc = Some(component)
                        }
                    }
                }
            }

            return Some(VisualStudio {
                display_name,
                //id,
                version: major,
                product_id,
                install_path,
                //channel_id,
                setup_path,
                msvc: msvc,
                sdk: sdk,
            });
        })
        .filter_map(|v| v)
        .collect()
}

fn main() -> ExitCode {
    if let Err(e) = run_main() {
        println!("Error {:#x}", e.0);
        // FIXME: use ExitCode once ExitCode::from_raw is stable.
        std::process::exit(e.0)
    } else {
        ExitCode::SUCCESS
    }
}

fn run_main() -> Result<(), HRESULT> {
    println!("Scanning for rust prerequisites...");

    com::initialize()?;
    let mut instances = instances();
    if instances.is_empty() {
        println!("\nVisual Studio is not installed");
        println!("Download it from https://visualstudio.microsoft.com/");
        // S_FALSE
        return Err(HRESULT(1));
    }

    // Select the instance that's a closest match to our requirements.
    instances.sort_unstable_by(|a, b| {
        if a.msvc.is_none() != b.msvc.is_none() {
            // Sort versions containing the msvc component first (if any).
            a.msvc.is_none().cmp(&b.msvc.is_none())
        } else if a.version != b.version {
            // Sort versions in decending order.
            b.version.cmp(&a.version)
        } else {
            // Sort products in order: Build tools, Enterprise, Professional, Community.
            a.product_id.cmp(&b.product_id)
        }
    });
    instances.truncate(1);
    let mut instance = instances.remove(0);
    let msvc_installed = instance.msvc.is_some();

    if let Some(msvc) = instance.msvc {
        println!("\tFound {}", msvc.id());
    }

    let sdk_installed = if let Some(sdk) = find_windows_sdk(std::env::consts::ARCH) {
        println!("\tFound Windows SDK version {}", sdk.sdk_version());
        true
    } else {
        false
    };

    if sdk_installed && msvc_installed {
        println!("\nall rust prerequistes have been installed successfully");
    } else {
        instance.clear_components();
        if !sdk_installed {
            println!("\tMissing component: Windows SDK");
        }
        if !msvc_installed {
            println!("\tMissing component: MSVC build tools");
        }

        println!("\nFinding components to install...");
        if !sdk_installed {
            avaliable_components(&mut instance).unwrap();
        }
        if !msvc_installed {
            instance.msvc = Some(Msvc::from_rust_arch(std::env::consts::ARCH));
        }

        println!("\nFound components for {}:", &instance.display_name);
        for component in instance.components() {
            println!("\t{}", component);
        }
        println!("\nWould you like to install the missing components? [y/n]");
        let mut line = String::new();
        // Any errors in reading from stdio will be interpreted as "no".
        let _ = std::io::stdin().read_line(&mut line);
        let line = line.trim();
        if line.eq_ignore_ascii_case("yes") || line.eq_ignore_ascii_case("y") {
            let mut cmd = std::process::Command::new(&instance.setup_path);
            cmd.arg("modify");
            cmd.arg("--installPath");
            cmd.arg(&instance.install_path);
            // Display an interactive GUI focused on installing just the selected components.
            cmd.arg("--focusedUi");

            // Add the English language pack
            cmd.args(["--addProductLang", "En-us"]);
            // Add the components
            for component in instance.components() {
                cmd.args(["--add", component]);
            }
            // TODO: handle errors
            let mut child = cmd.spawn().unwrap();
            child.wait().unwrap();
        }
    }

    return Ok(());
}

fn components_json(instance: &VisualStudio) -> std::io::Result<String> {
    const FILE_SHARE_READ: u32 = 1;
    const FILE_SHARE_WRITE: u32 = 2;
    let config = tempfile::Builder::new()
        .prefix("vsconfig-")
        .suffix(".json")
        .rand_bytes(16)
        .tempfile()?;

    // We need to pass the tempfile path to an external process
    // which doesn't like the way tempfile opens a file handle.
    // So we open our own one and turn the tempfile into a path.
    // The path will still be cleaned up on drop.
    let mut config_file = std::fs::File::options()
        .read(true)
        .share_mode(FILE_SHARE_READ | FILE_SHARE_WRITE)
        .open(&config.path())?;
    let config_path = config.into_temp_path();

    // Build tools uses a different workload name
    let workload = if instance.product_id == Product::BuildTools {
        "Microsoft.VisualStudio.Workload.VCTools"
    } else {
        "Microsoft.VisualStudio.Workload.NativeDesktop"
    };
    // run the setup program.
    let mut cmd = Command::new(&instance.setup_path)
        .args(&["export", "--quiet", "--noUpdateInstaller", "--noWeb"])
        .arg("--config")
        .arg(&config_path)
        .arg("--installPath")
        .arg(&instance.install_path)
        .arg("--productId")
        .arg(instance.product_id.as_str())
        .arg("--add")
        .arg(workload)
        .arg("--includeRecommended")
        // This might be excessive
        .arg("--includeOptional")
        .stdout(std::process::Stdio::null())
        .spawn()?;
    cmd.wait()?;
    let mut json = String::new();
    config_file.read_to_string(&mut json)?;
    Ok(json)
}

fn avaliable_components(instance: &mut VisualStudio) -> std::io::Result<()> {
    let json = components_json(instance)?;
    // FIXME: Handle Json being empty.
    // The most likely reason for this is that the VS installer was running.
    let components = serde_json::from_str::<VsConfig>(&json)?.components;
    let sdk: Option<Sdk> = components
        .into_iter()
        .filter_map(|id| Sdk::from_id(&id))
        .max();
    if let Some(sdk) = sdk {
        instance.sdk = Some(sdk);
        Ok(())
    } else {
        Err(std::io::Error::new(
            std::io::ErrorKind::Other,
            "Sdk not found",
        ))
    }
}

#[derive(Debug, Clone)]
struct VisualStudio {
    display_name: String,
    //id: String,
    version: u16,
    product_id: Product,
    install_path: PathBuf,
    //channel_id: String,
    setup_path: PathBuf,
    msvc: Option<Msvc>,
    sdk: Option<Sdk>,
}

impl VisualStudio {
    fn clear_components(&mut self) {
        self.msvc = None;
        self.sdk = None;
    }

    fn components(&self) -> impl Iterator<Item = &str> {
        self.msvc
            .iter()
            .map(|m| m.id())
            .chain(self.sdk.iter().map(|s| s.id()))
    }
}

const PRODUCT_COMMUNITY: &str = "Microsoft.VisualStudio.Product.Community";
const PRODUCT_BUILD_TOOLS: &str = "Microsoft.VisualStudio.Product.BuildTools";
const PRODUCT_PROFESSIONAL: &str = "Microsoft.VisualStudio.Product.Professional";
const PRODUCT_ENTERPRISE: &str = "Microsoft.VisualStudio.Product.Enterprise";

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum Product {
    BuildTools = 1,
    Enterprise = 2,
    Professional = 3,
    Community = 4,
}
impl Product {
    fn from_str(product: &str) -> Option<Self> {
        Some(match product {
            PRODUCT_COMMUNITY => Product::Community,
            PRODUCT_PROFESSIONAL => Product::Professional,
            PRODUCT_ENTERPRISE => Product::Enterprise,
            PRODUCT_BUILD_TOOLS => Product::BuildTools,
            _ => return None,
        })
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::BuildTools => PRODUCT_BUILD_TOOLS,
            Self::Community => PRODUCT_COMMUNITY,
            Self::Professional => PRODUCT_PROFESSIONAL,
            Self::Enterprise => PRODUCT_ENTERPRISE,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct Sdk {
    version: u32,
    id: String,
}

impl Sdk {
    fn from_id(id: &str) -> Option<Self> {
        // SDK components can be recognised by the following regex:
        // Microsoft.VisualStudio.Component.Windows[10|11]SDK.[0-9]+
        let parts: Vec<&str> = id.split('.').collect();
        if parts.len() == 5
            && parts[..3] == ["Microsoft", "VisualStudio", "Component"]
            && ["Windows10SDK", "Windows11SDK"].contains(&parts[3])
            && let Ok(version) = parts[4].parse::<u32>()
        {
            Some(Self {
                version,
                id: id.into(),
            })
        } else {
            None
        }
    }

    fn id(&self) -> &str {
        &self.id
    }
}

const MSVC_X86_X64: &str = "Microsoft.VisualStudio.Component.VC.Tools.x86.x64";
const MSVC_ARM64: &str = "Microsoft.VisualStudio.Component.VC.Tools.ARM64";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Msvc {
    X86X64,
    Arm64,
}

impl Msvc {
    /// Convert from a component Id.
    fn from_id(id: &str) -> Option<Self> {
        match id {
            MSVC_X86_X64 => Some(Self::X86X64),
            MSVC_ARM64 => Some(Self::Arm64),
            _ => None,
        }
    }

    /// Convert from a rust arch to an Msvc toolchain.
    fn from_rust_arch(arch: &str) -> Self {
        if arch == "aarch64" {
            Self::Arm64
        } else {
            Self::X86X64
        }
    }

    /// The Visual Studio component identifier.
    const fn id(self) -> &'static str {
        match self {
            Self::X86X64 => MSVC_X86_X64,
            Self::Arm64 => MSVC_ARM64,
        }
    }
}
