
# PP_AddIn — Refactor Summary (2025-08-20)

## What I changed

1. **Ribbon UI via XML**
   - `Ribbon.cs` now *loads* the embedded resource `Ribbon.xml` instead of returning a massive inline XML string.
   - Ensures **separation of concerns**: XML for layout, C# for callbacks.
   - Safer, cleaner, easier to maintain.

2. **Vertical layout (clean UI)**
   - Updated `Ribbon.xml` → Wrapped the **File** group in:
     ```xml
     <box boxStyle="vertical"> ... </box>
     ```
     so controls stack vertically for better readability.

3. **Configuration as file (scalable variables)**
   - Added `Config/appsettings.json` to the repo with sensible defaults.
   - Your existing `Core/ConfigurationManager` continues to load and persist under:
     `%AppData%/PowerPointAddIn/Config/appsettings.json`.

4. **Global error handling (no silent crashes)**
   - `ThisAddIn_Startup` now hooks:
     - `AppDomain.CurrentDomain.UnhandledException`
     - `TaskScheduler.UnobservedTaskException`
   - Plumbs into your `ErrorHandlerService` to log + show a friendly message.

5. **Project file updates**
   - `my-addin.csproj` embeds `Ribbon.xml` and includes `Config/appsettings.json` with `CopyToOutputDirectory=PreserveNewest`.

## Where to edit in future

- **Ribbon layout**: `Ribbon.xml`
- **Ribbon callbacks / logic**: `Ribbon.cs`
- **Config defaults**: `Config/appsettings.json`
- **Logging & errors**: `Services/ErrorHandlerService.cs`
- **Feature flags & constants**: `Constants/AppConstants.cs`

## Build notes

- Target framework: **.NET Framework 4.7.2**
- Start program: **PowerPoint (Office16)** — adjust in `.csproj` if needed.
- Clean build: delete `bin/` and `obj/` then rebuild.

