# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TailorMail (千函) is a WPF desktop app for personalized bulk email sending. Users compose email templates with variable placeholders (e.g. `{Name}`, `{ShortName}`), define per-recipient values, and send via Outlook COM or SMTP (MailKit). The UI is a 6-step wizard: recipients -> variables -> compose -> attachments -> preview -> send.

## Build & Run

Target: `net10.0-windows` (SDK 10.0.203, pinned in `global.json`). Windows-only, requires .NET 10 SDK.

```bash
# Restore
dotnet restore

# Build
dotnet build

# Run
dotnet run --project src/TailorMail/TailorMail.csproj

# Publish (self-contained single-file with trimming)
dotnet publish src/TailorMail/TailorMail.csproj -c Release -o publish
```

No test project exists. No linter configured.

## Architecture

### MVVM with CommunityToolkit.Mvvm

- **Models** (`Models/`): Plain C# classes. `Recipient`, `SendResult` use `ObservableObject` for binding.
- **ViewModels** (`ViewModels/`): Extend `ObservableObject`, use `[ObservableProperty]` and `[RelayCommand]`. Each View manually instantiates its ViewModel (no ViewModelLocator).
- **Views** (`Views/`): UserControls. MainWindow lazily creates and caches step pages (`_step1`–`_step6`). Pages implement `IRefreshable` for data reload on navigation.

### DI Container (App.xaml.cs)

```
IDataService        -> SqliteDataService   (singleton)
AttachmentMatchService                     (singleton)
OutlookEmailSender                         (singleton)
SmtpEmailSender                            (transient, IDisposable)
```

`App.Services` is the static `ServiceProvider`. `ServiceHelper.GetRequiredService<T>()` is the service locator for non-DI contexts.

### Email Sending

Two `IEmailSender` implementations:
- **OutlookEmailSender**: COM Interop. Must `Marshal.ReleaseComObject` in `finally` to prevent Outlook leaks. Runs in `Task.Run()`.
- **SmtpEmailSender**: MailKit. Transient, disposed after each send batch. Password passed at runtime (never stored in sender).

Send flow in `SendViewModel.ExecuteSend`: resolve sender by `SendMethod` -> convert FlowDocument XAML to HTML via `FlowDocumentHelper.ToHtml()` -> append signature -> iterate recipients with variable substitution -> delay between sends (`SendIntervalMs`) -> support cancellation and retry.

### Data Layer

SQLite database at `{AppDir}/data/tailormail.db`. Five tables:
- `RecipientGroups` / `Recipients`: Normalized relational. Recipients store `VariablesJson` as JSON blob.
- `AppSettings` / `AttachmentConfig`: Singleton-row JSON blob pattern (Id=1).
- `MailTemplates`: Template storage with XAML body.

`JsonDataService` exists only for legacy migration (writes `.migrated_to_sqlite` flag).

### Credential Storage

`CredentialHelper` uses Windows DPAPI (`ProtectedData`, `CurrentUser` scope). SMTP passwords encrypted before DB storage, decrypted at send time.

### Design System

`DesignTokens.cs` defines colors, typography, spacing (8px grid), radii, shadows, animation durations. All tokens available as both C# constants and XAML `StaticResource`s. UI uses WPF-UI (Wpf.Ui) controls: `Snackbar`, `InfoBar`, `SymbolIcon`.

## Key Patterns

- **Wizard navigation**: MainWindow caches pages, uses `Dispatcher.BeginInvoke` with `DispatcherPriority.Loaded` for transitions.
- **Debounced saving**: `RecipientsViewModel.ScheduleSave()` uses 500ms `DispatcherTimer`.
- **Rich text**: `MailComposePage` uses `RichTextBox` + `FlowDocumentHelper` for HTML/XAML conversion. Templates store XAML body.
- **Attachment matching**: `AttachmentMatchService` auto-matches files by recipient name/shortname.
- **Variable substitution**: Built-in `{Name}`, `{ShortName}` + custom variables. `ProcessBody()` replaces all placeholders.
- **Friendly errors**: `SendResult.GetFriendlyError()` maps SMTP errors to user-readable Chinese messages.
