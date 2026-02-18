Query Visual Studio setup for information on installed instances of Visual Studio.

This is a thin wrapper around the COM interface.
Consult the [`Microsoft.VisualStudio.Setup.Configuration`] documentation for more information on the API.

## Example

```rust
use vssetup::{com, HRESULT, SetupConfiguration};

fn main() -> Result<(), HRESULT> {
    com::initialize();
    let setup = vsinstance::SetupConfiguration::new()?;
    let instances = setup.EnumAllInstances()?;
    for instance in instances {
        let name = instance.GetDisplayName(0x400)?.to_string();
        println!("{name}");
    }
    Ok(())
}
```

[`Microsoft.VisualStudio.Setup.Configuration`]: https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.setup.configuration
