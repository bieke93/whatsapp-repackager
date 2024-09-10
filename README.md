# WhatsApp Repackager

WhatsApp Repackager is a simple tool designed to repackage WhatsApp conversations exported as a zip file into a more human-readable and preservation-friendly format. It extracts and reorganizes the contents of the zip file, outputs the data in preservation-friendly formats, and provides basic statistics along with a simple visualization of the conversation in an Excel file. For privacy reasons, participant names can be pseudonymized. Additionally, textual descriptions of emojis can be added to ensure they are understandable across different systems and over time.

This tool is a work-in-progress, and we welcome all feedback and suggestions.

## Getting Started

### One-Time Setup:
1. **Download and Open the Code**: Clone or download the repository.
2. **(Optional) Set Your Preferences**: You can configure key parameters at the top of the Python script (or skip this step and set them at runtime).
3. **Install Dependencies**: Run `pip install -r requirements.txt` to install the required Python packages.

### Repackaging a WhatsApp Zip File:
1. **Run the Script**: Execute `python whatsapp-repackager.py` (adjust as needed based on your Python installation) from anywhere in your system. The script will prompt you to provide the path to your WhatsApp zip file.

## Key Features

- **Zip File Extraction**: The tool extracts the contents of the provided zip file.
- **Conversation Parsing**: It identifies and processes the `.txt` conversation file, ensuring each line contains a complete message.
- **CSV, JSON, and Excel Output**: The tool generates output in CSV, JSON, and/or Excel formats, based on user selection. The data includes:
  - Conversation name
  - Unique ID for each message
  - Timestamp (date and time) of the message
  - Links to message attachments (not included in JSON)
  - Message content
  - A column for each participant
  - In the Excel file, each participant's messages are highlighted with a unique color for easy reading.
- **Summary File**: A separate CSV file (or worksheet in Excel) with basic conversation statistics is created.
- **Attachment Organization**: Attachments are organized into subfolders within an `attachments` directory, named after their respective message IDs.
- **Slugified Attachment Names**: Attachment filenames are transformed into a "slugified" format for consistency and ease of use.
- **Emoji Descriptions**: Optionally, emoji descriptions (e.g., 😊 -> [smiling face with smiling eyes]) are added to messages to ensure interpretability over time, using the [Emoji API](https://emoji-api.com/).
- **Pseudonymization**: The tool can pseudonymize participant names, generating a shortuuid for each participant. A mapping of the pseudonyms is saved to a CSV file, and the pseudonyms are applied to all output files, including a pseudonymized copy of the original `.txt` file.

## Limitations

- WhatsApp exports do not contain:
  - The creation date of the conversation
  - Comprehensive participant identities (only locally stored info is available)
  - Details on when participants joined or left the conversation
  - Information on message replies or threads
  - Attachment metadata (which is removed when sent by WhatsApp)

## License

This project is licensed under the [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/) license.

## AI Assistance

This project was developed with assistance from AI tools.
