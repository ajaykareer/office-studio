# Changelog

## 1.3.0

- Added explicit preparation, download, file-check, and ready stages driven by worker events.
- Added reassuring messages for long downloads, saved-file reuse, and unchanged file sizes while Microsoft setup is still responding.
- Kept elapsed time and live MB visible without an invented ETA or overall percentage.
- Added a Windows ready notification, background taskbar attention, and a tray icon that returns to the app without starting installation.
- Highlighted Install Office when the package is ready, while preserving license consent.
- Added stage-order, progress, cancellation-notification, and native notification checks.

## 1.2.2

- Fixed a startup crash caused by updating the file-size label from a background thread before the WPF event loop was running.
- Added an asynchronous startup regression check.
- Renamed the executable and main window to **Office Studio 1.2.2**, retaining KAREER branding inside.

## 1.2.1

- Simplified the action button to **Download** in every state.
- Added a visible existing-files size in MB.
- This version introduced a startup crash; use 1.2.2 or later.

## 1.2.0

- Added Abort download, scoped to the download worker and its child processes.
- Added reuse of matching partial downloads and verified completed packages.
- Updated sidebar steps during download and installation readiness.

## 1.1.1

- Refreshed file metadata when displaying download file sizes.
- Clarified that files on disk may include reserved space.

## 1.1.0

- Added installed Office detection, compatibility checks, and removal controls.
- Added standard uninstall and Microsoft cleanup-tool integration.

## 1.0.0

- Initial desktop launcher with Office edition selection, app checkboxes, activity logs, and separate download/install controls.
