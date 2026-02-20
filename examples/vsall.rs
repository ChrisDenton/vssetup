//! List all completed Visual Studio installs.
//!
//! Use `cargo run --example vsall` to compile and run this.
//! If you also want to list all packages then use the `--packages` arguments.
//! E.g. `cargo run --example vsall -- --packages`.

use vssetup::{HRESULT, SetupConfiguration, com};

fn main() -> Result<(), HRESULT> {
    com::initialize()?;

    let list_packages = std::env::args().skip(1).any(|arg| arg == "--packages");
    let display_help = std::env::args()
        .skip(1)
        .any(|arg| arg == "-h" || arg == "--help");
    if display_help {
        println!("usage: vsall [--packages]");
        return Ok(());
    }

    let setup = SetupConfiguration::new()?;
    let mut first = false;
    for instance in setup.EnumInstances()? {
        if first {
            first = false;
        } else {
            println!();
        }
        println!(
            "displayName: {}",
            instance.GetDisplayName(0x400)?.to_string()
        );
        println!(
            "description: {}",
            instance.GetDescription(0x400)?.to_string()
        );
        println!("instanceId: {}", instance.GetInstanceId()?);
        println!(
            "installDate: FILETIME({})",
            &instance.GetInstallDate()?.as_u64()
        );
        println!("installationPath: {}", instance.GetInstallationPath()?);
        println!(
            "installationVersion: {}",
            instance.GetInstallationVersion()?
        );
        println!("state: {}", instance.GetState()?);
        println!("enginePath: {}", instance.GetEnginePath()?);
        println!("productPath: {}", instance.GetProductPath()?.to_string());
        if let Ok(Some(product)) = instance.GetProduct() {
            println!("product: {{");
            println!("    id: {}", product.GetId()?);
            println!("    uniqueId: {}", product.GetUniqueId()?);
            println!("    version: {}", product.GetVersion()?);
            println!("    type: {}", product.GetType()?);
            println!("    branch: {}", product.GetBranch()?);
            println!("    chip: {}", product.GetChip()?);
            println!("    isExtension: {}", product.GetIsExtension()?);
            println!("    isInstalled: {}", product.GetIsInstalled()?);
            println!("    language: {}", product.GetLanguage()?);
            println!(
                "    supportsExtensions: {}",
                product.GetSupportsExtensions()?
            );
            println!("}}");
        }
        if let Ok(properties) = instance.to_property_store() {
            println!("propertyStore: {{");
            for property in properties.GetNames()?.iter() {
                let value = properties.GetValue(property)?;
                println!("    {property}: {value}");
            }
            println!("}}");
        }
        if let Ok(Some(properties)) = instance.GetProperties() {
            println!("properties: {{");
            for property in properties.GetNames()?.iter() {
                let value = properties.GetValue(property)?;
                println!("    {property}: {value}");
            }
            println!("}}");
        }

        if let Ok(catalog) = instance.to_catalog() {
            println!("catalog: {{");
            println!("    isPrerelease: {}", catalog.IsPrerelease()?);
            if let Ok(Some(properties)) = catalog.GetCatalogInfo() {
                for property in properties.GetNames()?.iter() {
                    let value = properties.GetValue(property)?;
                    println!("    {property}: {value}");
                }
            }
            println!("}}");
        }
        if list_packages && let Ok(packages) = instance.GetPackages() {
            println!("packages: [");
            for package in packages.iter() {
                println!("    {{");
                println!("        id: {}", package.GetId()?);
                println!("        uniqueId: {}", package.GetUniqueId()?);
                println!("        version: {}", package.GetVersion()?);
                println!("        type: {}", package.GetType()?);
                println!("        branch: {}", package.GetBranch()?);
                println!("        chip: {}", package.GetChip()?);
                println!("        isExtension: {}", package.GetIsExtension()?);
                println!("        language: {}", package.GetLanguage()?);
                println!("        }}");
            }
            println!("]");
        }
    }
    Ok(())
}
