# MistHelper

MistHelper is a comprehensive Python CLI tool for interacting with the Juniper Mist API. It enables you to export, analyze, and manage data from your Mist organization, including sites, devices, alarms, events, and more. The tool is designed to minimize API usage by caching data in CSV files and supports both interactive and command-line modes.

## Features

- Export organization alarms, device events, audit logs, and inventory to CSV
- Fetch and display device and site statistics
- Export event and alarm definitions
- Enrich device and gateway data with site/location info
- Generate support packages for troubleshooting
- Interactive CLI for device shell access and command execution
- Rate-limited API calls with dynamic delay to avoid exceeding Mist API limits
- Multi-threaded fast mode for bulk data collection
- Data flattening and CSV-friendly formatting

## Requirements

- Python 3.7+
- The following Python packages (auto-installed if missing):
  - mistapi
  - websocket-client
  - pyte
  - requests
  - prettytable
  - tqdm
  - sshkeyboard
  - numpy
  - python-dotenv

## Setup

1. Clone this repository:
   ```sh
   git clone https://github.com/yourusername/MistHelper.git
   cd MistHelper
   ```
2. Create a `.env` file in the project directory with your Mist API credentials:
   ```ini
   MIST_HOST=api.mist.com
   MIST_APITOKEN=your_api_token_here
   org_id=your_org_id_here
   ```
3. Run the script. Required packages will be installed automatically if missing.

## Usage

You can run MistHelper in two modes:

### Interactive Menu

Simply run:
```sh
python MistHelper.py
```
You will be presented with a menu of available actions.

### Command-Line Arguments

You can also run specific actions directly:
```sh
python MistHelper.py --menu 1 --org your_org_id
```

#### Common Arguments
- `-O`, `--org` : Organization ID
- `-M`, `--menu` : Menu option number to execute (see below)
- `-S`, `--site` : Human-readable site name
- `-D`, `--device` : Human-readable device name
- `-P`, `--port` : Port ID
- `--debug` : Enable debug output
- `--delay` : Fixed delay between loop iterations (in seconds)
- `--fast` : Enable fast mode with multithreading

## Menu Options

| Option | Description |
|--------|-------------|
| 0      | Select a site (used by other functions) |
| 1      | Export all organization alarms from the past day |
| 2      | Export all device events from the past 24 hours |
| 3      | Export audit logs for the organization |
| ...    | ... (see script for full list) |

For a full list of options, run the script without arguments.

## Example: Export All Sites
```sh
python MistHelper.py --menu 11 --org your_org_id
```

## Data Caching

MistHelper caches API responses in CSV files. If a CSV is fresh (default: 15 minutes), it will be used instead of making a new API call. This helps conserve API requests and avoid rate limits.

## Support & Contributions

- Issues and pull requests are welcome!
- Please open an issue for bugs or feature requests.

## License

MIT License
