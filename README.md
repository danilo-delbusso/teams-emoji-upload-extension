# Teams Emoji Uploader

A Chrome extension for managing your organisation's custom emojis in Microsoft Teams on the web. It opens as a side panel next to Teams, where you can upload emojis in bulk, rename them before they go up, and browse or delete existing ones.

## Features

### Upload

- Add many images at once. Each one gets a preview and an editable `:name:`, which defaults to the file name.
- Uploads up to 10 files in parallel and shows the progress of each one.
- When an upload fails, the reason is shown on that file (for example, the name is already taken), and you can rename it and retry just the failed ones.
- The file list survives closing the panel, and an upload keeps running in the background if you do.

### Browse and delete

- Search your organisation's custom emojis, or show only the ones you added.
- Delete one emoji or a whole selection. Every delete asks for confirmation and lists what will be removed; deleting 10 or more also requires typing `delete`.
- Teams allows each user 10 deletions per minute. Larger deletes are paced to stay within that, and pause and retry automatically if Teams asks them to slow down.

### Other

- After uploading, the extension can refresh Teams so new emojis show up: either by clearing Teams' caches, which keeps you signed in, or by a full reset, which signs you out.
- Light and dark themes in Teams' default colours, following your system setting until you pick one.

## Requirements

- Chrome 114 or later, or another Chromium-based browser that supports extension side panels.
- Teams open in the browser at `teams.microsoft.com` or `teams.cloud.microsoft`, and signed in.
- Permission to add custom emojis in your organisation. Teams may only let you delete emojis you added yourself.

## Installation

### Chrome Web Store

[Teams Emoji Uploader](https://chromewebstore.google.com/detail/teams-emoji-uploader/mlajagdepghhbclefnmcdnjfhfdmoofo) is the published version of the upstream project. It may not have every feature described here yet.

### From source

1. Clone this repository.
2. Install dependencies and build:

   ```console
   npm install
   npm run build
   ```

3. Open `chrome://extensions` and turn on **Developer mode**.
4. Click **Load unpacked** and select the repository folder.

   ![Load unpacked](./readme/unpack.png)

5. Open Teams in the browser and click the extension's toolbar icon to open the side panel.

After changing the code, run `npm run build` again and click the reload icon on the extension's card in `chrome://extensions`.

## Usage

1. Open Teams in the browser and open the side panel.
2. **Upload**: click **Add Emoji Files**, adjust the names if needed, and click **Upload to Teams**.
3. **Browse & delete**: search or filter the list, then use the trash icon on an emoji, or select several and click **Delete selected**.
4. If new or deleted emojis don't show up in Teams' emoji picker, click **Refresh Teams**. Teams can take a while to sync changes on its side.

## How it works

The extension uses the same internal Teams web APIs that Teams itself calls. They are undocumented, so they may change without notice.

- **Sign-in**: it reads the Teams access tokens from the Teams tab's local storage and keeps them in the extension's local storage while it works. Tokens are only sent to Microsoft's Teams endpoints.
- **Upload**: the image is stored through Teams' file service (`asyncgw.teams.microsoft.com`), then registered under its name with `POST /api/csa/{region}/api/v1/customemoji/metadata`.
- **Browse**: `GET /api/csa/{region}/api/v1/customemoji/metadata` lists the custom emojis. Previews are loaded from the file service.
- **Delete**: `DELETE /api/csa/{region}/api/v1/customemoji/metadata/{name};{documentId}` removes the emoji's name. Like deleting in Teams itself, the stored image is not removed.

### Permissions

| Permission                           | Used for                                                      |
| ------------------------------------ | ------------------------------------------------------------- |
| `activeTab`, `scripting`             | Reading the Teams sign-in tokens from the Teams tab           |
| `storage`, `unlimitedStorage`        | Keeping the file list and settings while the panel is closed  |
| `sidePanel`                          | Showing the extension next to Teams                           |
| `browsingData`                       | Refreshing or resetting Teams' site data so new emojis appear |
| Teams and `asyncgw` host permissions | Calling the Teams APIs listed above                           |

## Development

```console
npm run build         # build into dist/
npm test              # run unit tests with coverage
npm run format        # format with Prettier
npm run format:check  # check formatting, as CI does
npm run zip           # build and package chrome-extension.zip
```

Source layout:

| File                                     | Purpose                                                     |
| ---------------------------------------- | ----------------------------------------------------------- |
| `popup.html`, `src/popup.ts`             | Side panel UI: upload list, settings, theme, view switching |
| `src/emojiBrowser.ts`                    | Browse and delete view, including the confirmation dialog   |
| `src/background.ts`                      | Service worker that runs uploads and reports progress       |
| `src/msTeams.ts`                         | Client for the Teams APIs                                   |
| `src/rateLimit.ts`                       | Handling Teams' 429 responses and pacing requests           |
| `src/draft.ts`                           | Saving the file list so it survives closing the panel       |
| `src/shortcuts.ts`, `src/customEmoji.ts` | Name defaults and validation, emoji list filtering          |

## Contributing

Contributions are welcome. Please open a pull request.

## License

[MIT License](LICENSE)
