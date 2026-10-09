# How it works

The extension uses the same internal Teams web APIs that Teams itself calls. They are undocumented, so they may change without notice.

- **Sign-in**: it reads the Teams access tokens from the Teams tab's local storage and keeps them in the extension's local storage while it works. Tokens are only sent to Microsoft's Teams endpoints.
- **Upload**: the image is stored through Teams' file service (`asyncgw.teams.microsoft.com`), then registered under its name with `POST /api/csa/{region}/api/v1/customemoji/metadata`.
- **Browse**: `GET /api/csa/{region}/api/v1/customemoji/metadata` lists the custom emojis. Previews are loaded from the file service.
- **Delete**: `DELETE /api/csa/{region}/api/v1/customemoji/metadata/{name};{documentId}` removes the emoji's name. Like deleting in Teams itself, the stored image is not removed.
- **Rate limits**: Teams answers `429 Too Many Requests` with the limit and how long to wait (currently 10 deletions per user per minute). The extension paces requests to stay under it and retries after the wait.

## Permissions

| Permission                           | Used for                                                        |
| ------------------------------------ | --------------------------------------------------------------- |
| `activeTab`, `scripting`             | Reading the Teams sign-in tokens from the Teams tab             |
| `storage`, `unlimitedStorage`        | Keeping the file list and upload progress when the panel closes |
| `sidePanel`                          | Showing the extension next to Teams                             |
| `browsingData`                       | Refreshing or resetting Teams' site data so new emojis appear   |
| Teams and `asyncgw` host permissions | Calling the Teams APIs listed above                             |
