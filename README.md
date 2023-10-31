## Goal:

This attempts to insert the contents of another DOCX file into the current DOCX file, by creating a "content control" in the current DOCX file, then use **`cc.insertFileFrombase64`** to insert a `base64`.

## Unfortunately:

It does not work reliably. Calling the `.insertFileFromBase64` function breaks the MS Word editor app UI.

The editor freezes, shows a "Waiting..." or "Inserting...", but the dialog never goes away until the page is refreshed.

<img width="1270" alt="image" src="https://github.com/paulshorey/word-editor-poc/assets/7524065/6e31df33-21ef-4995-8430-22ff51befaa2">

## STRANGE THING IS... THIS IS A CLUE...

Loading the same OneDrive/SharePoint document in a different tab or even a different computer will show the inserted content correctly - and even triggers the original "frozen" Microsoft Word editor tab to get unstuck and start working correctly again.

## Run the Word add-in:

- Open MS Word online (Sharepoint or OneDrive)

- Click the "add-ins" button in the top toolbar

- Upload the add-in manifest (specially hosted just for this debugging experiment): https://base64-word-editor-poc-debug.paulshorey.com/manifest.xml

- Click the "FIRST AMERICAN" add-in button in the toolbar

- Please keep in mind - it is intermittent. Sometimes it works. Sometimes it even works most of the time or almost all the time.

- (1) Place the cursor inside the word document

- (2) Click the “Select Component to Add” dropdown and pick an option.

- (3) Click "Add"

- Alternatively, enter your own "base64" or "OOXML" string into the text area below the select UI, then click the correct text-button below the text area.

## Debug/engineer add-in:

- `npm install` (or `yarn`)

- `npm run start` (or `yarn start`)

- open `https://localhost:3000` in a browser. Click anywhere in the page. Type "thisisunsafe"

- then follow the "Run the Word add-in" instructions above, but with the ROOT `manifest.xml` file instead of the one hosted on the server

## Debugging:

Please keep in mind - it is intermittent. Sometimes it works. Sometimes it even works most of the time or almost all the time.

We've tried different simple DOCX files converted to base64 or XML, tried to extract the `./word/document.xml` from the DOCX file.

We can use OOXML instead of base64, but getting the same error.

We even tried using the API to output the contents of a "content control" as OOXML, then use that exact correct OOXML with **`cc.insertOoxml`** to insert the file. But no bueno.

## Resources:

Convert word files to base64 string:
https://products.aspose.app/pdf/conversion/docx-to-base64
