# Changelog

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
