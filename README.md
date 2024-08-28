# Whatsapp-repackager

This script is a simple tool (first attempt) to repackage WhatsApp conversations exported as a zip file, in a human-readable and preservation-friendly way. It extracts and reorganises the contents of the containerfile, offers basic statistics and a simple visualisation of the conversation in an Excel-file.

## How to Make It Work

1. **Download and Open the Code**: Clone or download the repository containing this script.
2. **Get a Free API Key**: Obtain a free API key from [Emoji API](https://emoji-api.com/). Add it in the upper part of the script.
3. **Set the Export Language**: Ensure that the language of your WhatsApp export (based on the App language of who made the export) is set correctly. This is necessary for identifying references to deleted messages and attachments in messages.
4. **Install Required Packages**: If not already installed, make sure to install the following Python packages:
   - `pandas`
   - `requests`
   - `openpyxl`
   - `python-slugify`
5. **Run the Script**: You can run the script from anywhere. It will prompt you to provide the path to your WhatsApp zip file.

## What It Does

- **Extracts the Zip File**: The script extracts the contents of the provided zip file.
- **Finds and Parses the Conversation File**: It locates the `.txt` file containing the conversation (based on the file name) and adjusts it to ensure each line corresponds to a complete message.
- **Creates CSV and Excel Files**: It generates CSV and Excel files with the following columns:
  - Conversation name
  - Unique ID for each message
  - Date and time of the message
  - A link to the location of attachments referenced in the message
  - The message content
  - A specific column for each participant in the conversation

  In the Excel file, each participant's messages are highlighted in a unique color, making the conversation easier to read.

- **Organizes Attachments**: The script creates an `attachments` folder, with subfolders for each message (named after their ID). The attachment(s) in each message is/are moved to their respective subfolders.

## Other Information

- **Slugified Attachment Names**: Attachment filenames are slugified in both the CSV/Excel files and the actual filenames.
- **Emoji Descriptions**: Emoji in the messages are supplemented with their descriptions, enclosed in square brackets. This is achieved using [Emoji API](https://emoji-api.com/).

## Limitations

- A WhatsApp export does not include the following details about the conversation:
  - Creation date of the conversation
  - Identity of the participants beyond locally stored information
  - Details about who joined the conversation and when
  - Information about when messages are replies to other messages

## To Do

- Add an anonymization option
- Add an option to include additional contact details
- Add an option of integrating several zip's

## License

[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)

## AI Notice

AI assistance was used during the development of this script.
