import mistapi, csv, ast, json, time, logging, os, argparse, sys, websocket, threading, re, sys, shutil, pyte
from prettytable import PrettyTable
from tqdm import tqdm
from datetime import datetime, timedelta
from sshkeyboard import listen_keyboard, stop_listening

logging.basicConfig(
    filename='script.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

current_epoch = int(time.time()) # Get the current epoch timestamp
past_epoch = current_epoch - 24 * 3600 # 24 hours * 3600 seconds/hour

device_type=str("all")
csv_file="stuff.csv"
fields = ["key", "display", "description"]

# Initialize API session with environment file
apisession = mistapi.APISession(env_file=".env",console_log_level=20,logging_log_level=20)
apisession.login()

org_id=None

print(mistapi.api.v1.self.usage.getSelfApiUsage(apisession).data)

def check_and_generate_csv(file_name, generate_function, freshness_minutes=15):
    """
    Checks if a CSV file exists and is fresh (modified within the last `freshness_minutes`).
    If not, it runs the `generate_function` to regenerate the file.
    """
    from datetime import datetime, timedelta
    import os

    # Check if the file already exists
    if os.path.exists(file_name):
        # Get the last modified time of the file
        file_mtime = datetime.fromtimestamp(os.path.getmtime(file_name))
        
        # Check if the file is still fresh
        if datetime.now() - file_mtime < timedelta(minutes=freshness_minutes):
            # Log that the cached file is being used
            logging.info(f"✅ Using cached {file_name} (fresh)")
            return
        else:
            # Log that the file is stale and will be regenerated
            logging.info(f"♻️ {file_name} is older than {freshness_minutes} minutes. Regenerating...")
    else:
        # Log that the file does not exist and will be generated
        logging.info(f"📄 {file_name} not found. Generating...")

    # Call the function to generate the file
    logging.info(f"🔄 Running {generate_function.__name__} to generate {file_name}...")
    generate_function()
    logging.info(f"✅ {file_name} generated or refreshed.")

def prepare_and_write_csv(data, filename, sort_key=None):
    """
    Flattens, sanitizes, optionally sorts, and writes data to a CSV file.
    """
    # Flatten nested dictionaries and lists
    data = flatten_all_nested_fields(data)
    
    # Escape multiline strings for CSV compatibility
    data = escape_multiline_strings(data)
    
    # Sort data by the specified key if provided
    if sort_key:
        data = sorted(data, key=lambda x: x.get(sort_key, ""))
    
    # Write the processed data to a CSV file
    write_data_to_csv(data, filename)

def display_pretty_table(data, fields=None, sortby=None):
    """
    Displays a PrettyTable from a list of dictionaries.
    """
    # Return early if there's no data to display
    if not data:
        return

    # Use provided fields or extract all unique keys
    fields = fields or get_all_unique_keys(data)

    # Initialize the PrettyTable with field names
    table = PrettyTable()
    table.field_names = fields

    # Set the sort column if it's valid
    if sortby and sortby in fields:
        table.sortby = sortby

    # Add each row of data to the table
    for item in data:
        row = [item.get(field, "") for field in fields]
        table.add_row(row)

    # Log the table as a string
    logging.info("\n" + table.get_string())

def interactive_device_action(fetch_function, filename, description, device_type="all"):
    """
    Prompts user to select a site and device, fetches data using the provided function,
    and writes the result to a CSV file.
    """
    # Prompt user to select a site
    site_id = prompt_user_to_select_site_id_from_csv()
    if not site_id:
        return

    # Prompt user to select a device at the selected site
    device_id = prompt_user_to_select_device_id(site_id, device_type=device_type)
    if not device_id:
        return

    # Log the action being performed
    logging.info(f"{description} for device ID: {device_id}")

    # Fetch data using the provided function
    stats = fetch_function(apisession, site_id, device_id).data

    # Flatten and sanitize the data
    stats = flatten_all_nested_fields([stats])
    stats = escape_multiline_strings(stats)

    # Write the data to a CSV file
    write_data_to_csv(stats, filename)

    # Display the data in a table
    display_pretty_table(stats)

def process_and_merge_csv_for_sfp_address():
    """
    Processes OrgDevicePortStats.csv and AllDevicesWithSiteInfo.csv to merge SFP transceiver info
    with site and device address/location, outputting a new merged CSV.
    Only ports with a non-empty transceiver model are included.
    """
    # Automatically generate missing files if needed
    if not os.path.exists('OrgDevicePortStats.csv'):
        print("⚠️ OrgDevicePortStats.csv not found. Generating it now...")
        logging.info("OrgDevicePortStats.csv not found. Generating it now...")
        export_org_device_port_stats()

    if not os.path.exists('AllDevicesWithSiteInfo.csv'):
        print("⚠️ AllDevicesWithSiteInfo.csv not found. Generating it now...")
        logging.info("AllDevicesWithSiteInfo.csv not found. Generating it now...")
        export_all_devices_with_site_info()

    # Load site and device info, keyed by MAC address
    with open('AllDevicesWithSiteInfo.csv', mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        site_info = {
            row['mac']: {
                'site_name': row.get('site_name', ''),
                'site_address': row.get('site_address', ''),
                'device_name': row.get('name', '')
            } for row in reader
        }

    # Merge with port stats, skipping rows with blank/null transceiver model
    merged_data = []
    with open('OrgDevicePortStats.csv', mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            mac = row.get('mac')
            transceiver_model = row.get('xcvr_model', '').strip()
            if mac in site_info and transceiver_model:
                merged_data.append({
                    'site_name': site_info[mac]['site_name'],
                    'site_address': site_info[mac]['site_address'],
                    'device_name': site_info[mac]['device_name'],
                    'port_id': row.get('port_id', ''),
                    'transceiver_part_number': row.get('xcvr_part_number', ''),
                    'transceiver_model': transceiver_model,
                    'transceiver_serial_number': row.get('xcvr_serial', '')
                })

    # Write output to new CSV
    output_file = 'MergedTransceiverData.csv'
    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        fieldnames = [
            'site_name', 'site_address', 'device_name', 'port_id',
            'transceiver_part_number', 'transceiver_model', 'transceiver_serial_number'
        ]
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(merged_data)

    print(f"✅ Merged data written to {output_file}")

def get_org_id():
    """
    Retrieves the organization ID (org_id) for API calls.
    Priority:
    1. Use global org_id if already set.
    2. Try to load from .env file (MIST_ORG_ID).
    3. Prompt user to select org via mistapi CLI.
    """
    global org_id
    if org_id:
        logging.info(f"✅ Using org_id from global variable: {org_id}")
        return org_id

    # Try to load from .env if not already set
    try:
        with open(".env", "r") as f:
            for line in f:
                if line.strip().startswith("MIST_ORG_ID="):
                    org_id = line.strip().split("=", 1)[1].strip().strip('"')
                    if org_id:
                        logging.info(f"✅ Loaded org_id from .env: {org_id}")
                        return org_id
    except FileNotFoundError:
        logging.warning("⚠️ .env file not found.")

    # Prompt if still not set
    logging.info("🔍 No org_id found in .env or CLI. Prompting user...")
    org_id_list = mistapi.cli.select_org(apisession)
    org_id = org_id_list[0]
    return org_id

def flatten_nested_dict(d, parent_key='', sep='_'):
    """
    Recursively flattens a nested dictionary, joining keys with `sep`.
    Lists of dicts are flattened with indexed keys.
    Non-dict lists are joined as comma-separated strings.
    """
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        # If the value is a dictionary, recurse
        if isinstance(v, dict):
            items.extend(flatten_nested_dict(v, new_key, sep=sep).items())
        # If the value is a list
        elif isinstance(v, list):
            if all(isinstance(i, dict) for i in v):
                # If all items are dicts, flatten each with an index
                for idx, item in enumerate(v):
                    items.extend(flatten_nested_dict(item, f"{new_key}{sep}{idx}", sep=sep).items())
            else:
                # Otherwise, join list items as a comma-separated string
                items.append((new_key, ','.join(map(str, v))))
        else:
            # Base case: not a dict or list, just add the value
            items.append((new_key, v))
    # Uncomment the next line to enable debug logging of the flattening process
    # logging.debug(f"Flattened dict at key '{parent_key}': {dict(items)}")
    return dict(items)

def flatten_all_nested_fields(data):
    """
    Flattens all nested fields in a list of dictionaries.
    - Attempts to parse stringified dicts/lists.
    - Recursively flattens nested dicts and lists of dicts.
    - Joins non-dict lists as comma-separated strings.
    """
    flattened = []
    for entry in data:
        new_entry = {}
        for key, value in entry.items():
            # Try to parse stringified dicts/lists
            if isinstance(value, str) and (value.startswith("{") or value.startswith("[")):
                try:
                    value = ast.literal_eval(value)
                    logging.debug(f"Parsed stringified value for key '{key}': {value}")
                except Exception:
                    try:
                        value = json.loads(value)
                        logging.debug(f"JSON loaded value for key '{key}': {value}")
                    except Exception:
                        # Leave as string if parsing fails
                        logging.debug(f"Failed to parse value for key '{key}', leaving as string.")

            # Flatten if it's a dict or list of dicts
            if isinstance(value, dict):
                # Recursively flatten nested dict
                flat = flatten_nested_dict(value, parent_key=key)
                new_entry.update(flat)
                logging.debug(f"Flattened dict for key '{key}': {flat}")
            elif isinstance(value, list):
                if all(isinstance(i, dict) for i in value):
                    # Flatten each dict in the list with an index
                    for idx, item in enumerate(value):
                        flat = flatten_nested_dict(item, parent_key=f"{key}_{idx}")
                        new_entry.update(flat)
                        logging.debug(f"Flattened dict in list for key '{key}_{idx}': {flat}")
                else:
                    # Join non-dict lists as comma-separated strings
                    new_entry[key] = ','.join(map(str, value))
                    logging.debug(f"Joined list for key '{key}': {new_entry[key]}")
            else:
                # Base case: not a dict or list, just add the value
                new_entry[key] = value
        flattened.append(new_entry)
    return flattened

def convert_list_values_to_strings(data):
    """
    Converts all list values in a list of dictionaries to comma-separated strings.
    Adds debug logging for each conversion.
    """
    for entry in data:
        for key, value in entry.items():
            if isinstance(value, list):
                # Log the conversion for debugging
                logging.debug(f"Converting list at key '{key}' to string: {value}")
                entry[key] = ','.join(map(str, value))
    return data

def get_all_unique_keys(data):
    """
    Returns a sorted list of all unique keys present in a list of dictionaries.
    Useful for determining CSV fieldnames or PrettyTable columns.
    """
    fields = set()
    for entry in data:
        # Add all keys from each dictionary to the set
        fields.update(entry.keys())
    # Log the discovered unique keys for debugging
    logging.debug(f"Discovered unique keys: {fields}")
    return sorted(fields)

def escape_multiline_strings(data):
    """
    Escapes multiline strings in a list of dictionaries for CSV compatibility.
    - Joins list values as comma-separated strings.
    - Replaces newline characters in strings with '\\n' and removes carriage returns.
    """
    for entry in data:
        for key, value in entry.items():
            if isinstance(value, list):
                # Convert list to comma-separated string for CSV compatibility
                logging.debug(f"Converting list at key '{key}' to string: {value}")
                entry[key] = ','.join(map(str, value))
            elif isinstance(value, str):
                # Replace newlines and carriage returns in strings
                if '\n' in value or '\r' in value:
                    logging.debug(f"Escaping newlines in string at key '{key}': {repr(value)}")
                entry[key] = value.replace('\n', '\\n').replace('\r', '')
    return data

def write_data_to_csv(data, csv_file):
    """
    Writes a list of dictionaries to a CSV file.
    - Escapes multiline strings for CSV compatibility.
    - Determines all unique fields for the CSV header.
    - Writes each row, filling missing fields with empty strings.
    """
    logging.debug(f"Preparing to write {len(data)} rows to {csv_file}...")
    data = escape_multiline_strings(data)
    fields = get_all_unique_keys(data)
    logging.debug(f"CSV fields determined: {fields}")
    with open(csv_file, 'w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        for idx, row in enumerate(data):
            # Write each row, ensuring all fields are present
            writer.writerow({field: row.get(field, "") for field in fields})
            if idx < 3:  # Log the first few rows for debugging
                logging.debug(f"Row {idx} written: {row}")
    logging.info(f"Data saved to {csv_file} ({len(data)} rows)")

def fetch_process_and_display_data(title, api_call, filename, sort_key=None, display_fields=None, **kwargs):
    """
    Fetches data using the provided API call, processes it (flattening, sorting, escaping),
    writes it to a CSV file, and displays it in a PrettyTable. Adds detailed logging.
    """
    logging.info(f"Starting data fetch: {title}")
    print(title)
    org_id = get_org_id()
    logging.debug(f"Using org_id: {org_id}")

    # Call the API and get all paginated results
    response = api_call(apisession, org_id, **kwargs)
    rawdata = mistapi.get_all(response=response, mist_session=apisession)
    logging.info(f"Fetched {len(rawdata)} raw records from API.")

    # Filter out non-dict entries (defensive)
    data = [entry for entry in rawdata if isinstance(entry, dict)]
    logging.debug(f"Filtered to {len(data)} dict records.")

    # Sort data if a sort key is provided
    if sort_key:
        data = sorted(data, key=lambda x: x.get(sort_key, ""))
        logging.debug(f"Data sorted by key: {sort_key}")

    # Flatten nested fields for CSV compatibility
    data = flatten_all_nested_fields(data)
    logging.debug("Flattened all nested fields.")

    # Escape multiline strings for CSV
    data = escape_multiline_strings(data)
    logging.debug("Escaped multiline strings.")

    # Determine all unique fields for CSV and table display
    fields = get_all_unique_keys(data)
    logging.debug(f"Unique fields for CSV/table: {fields}")

    # Write processed data to CSV
    write_data_to_csv(data, filename)
    logging.info(f"Data written to {filename} ({len(data)} rows).")

    # Prepare and display PrettyTable
    table = PrettyTable()
    table.field_names = display_fields if display_fields else fields
    table.valign = "t"
    for item in tqdm(data, desc="Processing", unit="record"):
        row = [item.get(field, "") for field in table.field_names]
        table.add_row(row)
    logging.info("\n" + table.get_string())

def prompt_user_to_select_device_id(site_id, device_type="all", csv_filename="SiteInventory.csv"):
    """
    Prompts the user to select a device by index or name from the device inventory at a given site.
    Returns the corresponding device ID, or None if not found.
    """
    # Fetch device inventory for the specified site and device type
    rawdata = mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id, type=device_type).data
    if not rawdata:
        print("No devices found for the selected site.")
        logging.warning(f"No devices found for site_id: {site_id} with device_type: {device_type}")
        return None

    # Sort, flatten, and sanitize the inventory data for display and CSV export
    inventory = sorted(rawdata, key=lambda x: x.get("model", ""))
    inventory = flatten_all_nested_fields(inventory)
    inventory = escape_multiline_strings(inventory)
    write_data_to_csv(inventory, csv_filename)
    logging.info(f"Device inventory for site_id {site_id} written to {csv_filename}")

    # Prepare PrettyTable for user selection
    table = PrettyTable()
    table.field_names = ["Index", "name", "mac", "model", "serial"]
    index_to_device = {}
    name_to_device = {}

    # Populate the table and lookup dictionaries
    for idx, item in enumerate(inventory):
        table.add_row([idx, item.get("name", ""), item.get("mac", ""), item.get("model", ""), item.get("serial", "")])
        index_to_device[idx] = item
        name_to_device[item.get("name", "")] = item

    print(table)
    logging.info("Displayed device selection table to user.")

    user_input = input("Enter the index or name of the device to view device: ").strip()
    logging.debug(f"User input for device selection: {user_input}")

    # Try index selection
    if user_input.isdigit():
        idx = int(user_input)
        if idx in index_to_device:
            device_id = index_to_device[idx].get("id")
            logging.info(f"User selected device by index: {idx} (device_id: {device_id})")
            return device_id
        else:
            logging.warning("❌ Invalid index.")
            return None

    # Try name selection
    if user_input in name_to_device:
        device_id = name_to_device[user_input].get("id")
        logging.info(f"User selected device by name: {user_input} (device_id: {device_id})")
        return device_id

    logging.warning("❌ Device not found by name or index.")
    return None

def show_site_device_inventory(site_id, device_type="all", csv_filename="SiteInventory.csv"):
    """
    Fetches and displays the device inventory for a given site.
    - site_id: The ID of the site to fetch inventory for.
    - device_type: The type of device to filter (default: "all").
    - csv_filename: The filename to write the inventory CSV to.
    """
    logging.info(f"Fetching device inventory for site_id={site_id}, device_type={device_type}")
    rawdata = mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id, type=device_type).data
    if not rawdata:
        print("No devices found for the selected site.")
        logging.warning(f"No devices found for site_id={site_id} with device_type={device_type}")
        return

    # Sort inventory by model for easier viewing
    inventory = sorted(rawdata, key=lambda x: x.get("model", ""))
    # Flatten nested fields for CSV and table compatibility
    inventory = flatten_all_nested_fields(inventory)
    # Escape multiline strings for CSV compatibility
    inventory = escape_multiline_strings(inventory)
    # Get all unique fields for CSV/table columns
    fields = get_all_unique_keys(inventory)
    # Write inventory to CSV
    write_data_to_csv(inventory, csv_filename)
    logging.info(f"Device inventory written to {csv_filename} ({len(inventory)} rows)")

    # Prepare PrettyTable for display
    table = PrettyTable()
    table.field_names = fields

    # Attempt to sort the table by 'model' if present
    if "model" in fields:
        try:
            table.sortby = "model"
        except Exception as e:
            logging.warning(f"⚠️ Could not sort table by 'model': {e}")

    # Add each device as a row in the table
    for item in inventory:
        row = [item.get(field, "") for field in fields]
        table.add_row(row)

    # Log the table output for reference
    logging.info("\n" + table.get_string())

def prompt_user_to_select_site_id_from_csv(csv_file="SiteList.csv"):
    """
    Prompts the user to select a site by index or name from SiteList.csv.
    Returns the corresponding site ID.
    """
    # Ensure the site list CSV is fresh or generate it if missing/stale
    check_and_generate_csv(csv_file, export_org_site_list)

    # Load the site list from CSV
    with open(csv_file, mode='r', encoding='utf-8') as file:
        reader = list(csv.DictReader(file))
        index_to_site = {i: row for i, row in enumerate(reader)}
        name_to_site = {row["name"]: row for row in reader if "name" in row}

    # Display available sites to the user
    print("\nAvailable Sites:")
    for idx, row in index_to_site.items():
        print(f"[{idx}] {row.get('name', 'Unnamed')}")

    user_input = input("\nEnter site index or name: ").strip()
    logging.debug(f"User input for site selection: {user_input}")

    # Try index selection
    if user_input.isdigit():
        idx = int(user_input)
        if idx in index_to_site:
            site_id = index_to_site[idx].get("id")
            print(f"✅ Selected site: {index_to_site[idx].get('name')} (ID: {site_id})")
            logging.info(f"User selected site by index: {idx} (site_id: {site_id})")
            return site_id
        else:
            print("❌ Invalid index.")
            logging.warning(f"Invalid site index entered: {idx}")
            return None

    # Try name selection
    if user_input in name_to_site:
        site_id = name_to_site[user_input].get("id")
        print(f"✅ Selected site: {user_input} (ID: {site_id})")
        logging.info(f"User selected site by name: {user_input} (site_id: {site_id})")
        return site_id

    print("❌ Site not found by name or index.")
    logging.warning(f"Site not found by name or index: {user_input}")
    return None

def select_site():
    """
    Prompts the user to select a site from the CSV list and logs the selection.
    """
    logging.info("Prompting user to select a site from SiteList.csv...")
    site_id = prompt_user_to_select_site_id_from_csv()
    if site_id:
        logging.info(f"✅ Selected site ID: {site_id}")
        # You can store or use the selected site_id as needed here
    else:
        logging.warning("❌ No site selected. User may have entered an invalid value or cancelled the prompt.")

def search_org_alarms():
    """
    Fetches all open organization alarms from the past 24 hours and writes them to OrgAlarms.csv.
    """
    logging.info("Starting search for all open org alarms in the past 24 hours...")
    fetch_process_and_display_data(
        title="Search all Org Alarms:",
        api_call=mistapi.api.v1.orgs.alarms.searchOrgAlarms,
        filename="OrgAlarms.csv",
        limit=1000,
        duration="24h",
        status="open"
    )
    logging.info("Completed search_org_alarms and wrote results to OrgAlarms.csv.")

def export_recent_device_events():
    """
    Export all device events from the past 24 hours to OrgDeviceEvents.csv.
    """
    logging.info("Search Org Device Events:")
    org_id = get_org_id()
    # Call the Mist API to search for device events in the last 24 hours
    response = mistapi.api.v1.orgs.devices.searchOrgDeviceEvents(
        apisession, org_id, device_type="all", limit=1000, last_by="-24h"
    )
    # Retrieve all paginated results
    rawdata = mistapi.get_all(response=response, mist_session=apisession)
    events = rawdata
    logging.info(f"Fetched {len(events)} device events from the past 24 hours.")
    # Write the events to a CSV file
    write_data_to_csv(events, "OrgDeviceEvents.csv")
    logging.info(f"Device events written to OrgDeviceEvents.csv ({len(events)} rows).")
    # Optionally log the first few events for debugging
    if events:
        logging.debug("Sample device events: %s", json.dumps(events[:3], indent=2))

def export_org_audit_logs():
    """
    Export organization audit logs to OrgAuditLogs.csv.
    Uses fetch_process_and_display_data to handle API call, CSV writing, and table display.
    """
    logging.info("Starting export of organization audit logs...")
    fetch_process_and_display_data(
        title="List Audit Logs:",
        api_call=mistapi.api.v1.orgs.logs.listOrgAuditLogs,
        filename="OrgAuditLogs.csv",
        limit=1000
    )
    logging.info("Completed export_org_audit_logs and wrote results to OrgAuditLogs.csv.")

def export_org_site_list():
    """
    Fetches and exports the list of all sites in the organization to SiteList.csv.
    Uses fetch_process_and_display_data to handle API call, CSV writing, and table display.
    """
    logging.info("Starting export of organization site list...")
    fetch_process_and_display_data(
        title="Site List:",
        api_call=mistapi.api.v1.orgs.sites.searchOrgSites,
        filename="SiteList.csv",
        sort_key="name",  # or "site_id" if preferred
        limit=1000
    )
    logging.info("Completed export_org_site_list and wrote results to SiteList.csv.")

def export_org_device_inventory():
    """
    Fetches and exports the full inventory of devices in the organization to OrgInventory.csv.
    Uses fetch_process_and_display_data to handle API call, CSV writing, and table display.
    """
    logging.info("Starting export of organization device inventory...")
    fetch_process_and_display_data(
        title="Org Inventory:",
        api_call=mistapi.api.v1.orgs.inventory.getOrgInventory,
        filename="OrgInventory.csv",
        sort_key="model",
        limit=1000
    )
    logging.info("Completed export_org_device_inventory and wrote results to OrgInventory.csv.")

def export_org_device_statistics():
    """
    Export statistics for all devices in the organization to OrgDeviceStats.csv.
    Uses fetch_process_and_display_data to handle API call, CSV writing, and table display.
    """
    logging.info("Starting export of organization device statistics...")  # Log start
    fetch_process_and_display_data(
        title="Org Device Stats:",
        api_call=mistapi.api.v1.orgs.stats.listOrgDevicesStats,
        filename="OrgDeviceStats.csv",
        sort_key="type",
        display_fields=[
            "name", "model", "mac", "status", "type", "hw_rev", "ip", "serial",
            "version", "height", "last_seen", "locating", "notes", "orientation", "uptime"
        ],
        type="all",
        limit=1000
    )

def export_org_device_port_stats():
    """
    Export port-level statistics for all switches and gateways in the organization to OrgDevicePortStats.csv.
    Uses fetch_process_and_display_data to handle API call, CSV writing, and table display.
    """
    logging.info("Starting export of organization device port statistics...")  # Log start of function
    fetch_process_and_display_data(
        title="Org Device Port Stats:",
        api_call=mistapi.api.v1.orgs.stats.searchOrgSwOrGwPorts,
        filename="OrgDevicePortStats.csv",
        sort_key="mac",
        display_fields=[
            "mac", "device_type", "device_interface_type", "speed",
            "rx_bytes", "tx_bytes", "port_desc", "port_id", "port_mac",
            "port_mode", "port_parent", "port_usage", "power_allocated", "power_draw"
        ],
        limit=1000
    )

def export_org_vpn_peer_stats():
    """
    Export VPN peer path statistics for the organization to OrgVPNPeerStats.csv.
    Uses fetch_process_and_display_data to handle API call, CSV writing, and table display.
    """
    logging.info("Starting export of organization VPN peer path statistics...")  # Log start of function
    fetch_process_and_display_data(
        title="Org VPN Peer Stats:",
        api_call=mistapi.api.v1.orgs.stats.searchOrgPeerPathStats,
        filename="OrgVPNPeerStats.csv",
        sort_key="mac",
        display_fields=[
            "hop_count", "is_active", "jitter", "latency", "loss", "mos",
            "network_interface", "peer_port_id", "peer_router_name", "port_id",
            "router_name", "up", "uptime", "vpn_name", "wan_name"
        ],
        limit=1000
    )

def interactive_view_site_inventory():
    """
    Prompts the user to select a site and displays its device inventory.
    """
    logging.info("Prompting user to select a site for device inventory view...")
    print("Select a Site to View Device Inventory:")
    site_id = prompt_user_to_select_site_id_from_csv()
    if site_id:
        logging.info(f"User selected site_id: {site_id} for inventory display.")
        show_site_device_inventory(site_id)
    else:
        logging.warning("No site selected or invalid input provided for site selection.")

def interactive_view_device_stats():
    """
    Prompts user to select a device and displays its detailed statistics.
    """
    logging.info("Prompting user to select a device for detailed statistics view...")
    # Call the interactive_device_action helper with the appropriate Mist API function
    interactive_device_action(
        fetch_function=mistapi.api.v1.sites.stats.getSiteDeviceStats,
        filename="DeviceStats.csv",
        description="Fetching detailed stats"
    )
    logging.info("Completed interactive_view_device_stats execution.")

def interactive_view_device_tests():
    """
    Prompts user to select a gateway device and displays its synthetic test stats.
    """
    logging.info("Prompting user to select a gateway device for synthetic test stats view...")
    # Call the interactive_device_action helper with the appropriate Mist API function
    interactive_device_action(
        fetch_function=mistapi.api.v1.sites.devices.getSiteDeviceSyntheticTest,
        filename="DeviceTestResults.csv",
        description="Fetching synthetic test stats",
        device_type="gateway"
    )
    logging.info("Completed interactive_view_device_tests execution.")

def interactive_view_device_config():
    """
    Prompts user to select a device and displays its configuration details.
    """
    logging.info("Prompting user to select a device for configuration details view...")  # Log start
    # Call the interactive_device_action helper with the appropriate Mist API function
    interactive_device_action(
        fetch_function=mistapi.api.v1.sites.devices.getSiteDevice,
        filename="DeviceConfig.csv",
        description="Fetching device configuration"
    )
    logging.info("Completed interactive_view_device_config execution.")  # Log completion

def export_all_org_devices():
    """
    Fetches and exports a list of all devices in the organization to OrgDevices.csv.
    Uses fetch_process_and_display_data to handle API call, CSV writing, and table display.
    """
    logging.info("Starting export of all organization devices...")  # Log start of function
    fetch_process_and_display_data(
        title="Org Devices:",
        api_call=mistapi.api.v1.orgs.devices.listOrgDevices,
        filename="OrgDevices.csv",
        sort_key="type",
        display_fields=["name", "mac"]
    )
    logging.info("Completed export_all_org_devices and wrote results to OrgDevices.csv.")  # Log completion

def fetch_all_site_settings(apisession, org_id, limit=1000):
    """
    Fetches configuration settings for all sites in the organization.

    Args:
        apisession: The Mist API session object.
        org_id: The organization ID.
        limit: (Unused) Maximum number of sites to fetch per API call.

    Returns:
        List of dictionaries, each containing the settings for a site.
    """
    logging.info("Fetching all site settings...")

    # Use mistapi.get_all to ensure pagination is handled for all sites
    response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id)
    sites = mistapi.get_all(response=response, mist_session=apisession)

    all_configs = []
    for site in tqdm(sites, desc="Sites", unit="site"):
        site_id = site.get("id")
        site_name = site.get("name", "Unnamed Site")
        try:
            # Fetch the site settings using the Mist API
            config = mistapi.api.v1.sites.setting.getSiteSetting(apisession, site_id).data
            config["site_id"] = site_id
            config["site_name"] = site_name
            all_configs.append(config)
            logging.info(f"✅ Fetched config for site: {site_name} (ID: {site_id})")
        except Exception as e:
            logging.warning(f"⚠️ Failed to fetch config for {site_name} (ID: {site_id}): {e}")

    logging.info(f"Fetched settings for {len(all_configs)} sites.")
    return all_configs

def export_all_site_settings():
    """
    Fetches and exports configuration settings for all sites in the organization to AllSiteConfigs.csv.
    Adds detailed logging at each step.
    """
    logging.info("Starting export of all site configuration settings...")  # Log start
    org_id = get_org_id()
    logging.debug(f"Using org_id: {org_id} for site settings export.")

    # Fetch all site settings using the helper function
    data = fetch_all_site_settings(apisession, org_id, limit=1000)
    if data:
        logging.info(f"Fetched settings for {len(data)} sites. Flattening and sanitizing data...")
        # Flatten nested fields for CSV compatibility
        data = flatten_all_nested_fields(data)
        # Escape multiline strings for CSV compatibility
        data = escape_multiline_strings(data)
        # Write the processed data to a CSV file
        write_data_to_csv(data, "AllSiteConfigs.csv")
        logging.info("✅ Site configs saved to AllSiteConfigs.csv")
    else:
        logging.warning("⚠️ No site configs found.")

def export_all_gateway_device_configs():
    """
    Fetches and exports configuration details for all gateway devices across all sites in the organization
    to AllSiteGatewayConfigs.csv. Adds detailed logging at each step.
    """
    logging.info("Starting export of all gateway device configurations...")  # Log start
    org_id = get_org_id()
    logging.debug(f"Using org_id: {org_id} for gateway device configs export.")

    # Fetch all gateway device configs using the helper function
    data = fetch_all_gateway_device_configs(apisession, org_id)
    if data:
        logging.info(f"Fetched configs for {len(data)} gateway devices. Flattening and sanitizing data...")
        # Flatten nested fields for CSV compatibility
        data = flatten_all_nested_fields(data)
        # Escape multiline strings for CSV compatibility
        data = escape_multiline_strings(data)
        # Write the processed data to a CSV file
        write_data_to_csv(data, "AllSiteGatewayConfigs.csv")
        logging.info("✅ Device configs saved to AllSiteGatewayConfigs.csv")
    else:
        logging.warning("⚠️ No device configs found.")

def fetch_all_gateway_device_configs(apisession, org_id):
    """
    Fetches configuration details for all gateway devices across all sites in the organization.

    Args:
        apisession: The Mist API session object.
        org_id: The organization ID.

    Returns:
        List of dictionaries, each containing the configuration for a gateway device.
    """
    logging.info("Fetching all sites in the org for gateway device configs...")

    # Use mistapi.get_all to paginate through all sites
    response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id, limit=1000)
    sites = mistapi.get_all(response=response, mist_session=apisession)
    logging.info(f"Found {len(sites)} sites in the organization.")

    all_device_configs = []

    for site in tqdm(sites, desc="Sites", unit="site"):
        site_id = site.get("id")
        site_name = site.get("name", "Unnamed Site")
        logging.debug(f"Processing site: {site_name} (ID: {site_id})")

        try:
            # Use mistapi.get_all to paginate through all gateway devices in the site
            response = mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id, type="gateway", limit=1000)
            devices = mistapi.get_all(response=response, mist_session=apisession)
            logging.info(f"  Found {len(devices)} gateway devices at site: {site_name} (ID: {site_id})")

            for device in tqdm(devices, desc=f"{site_name}", unit="device", leave=False):
                device_id = device.get("id")
                device_name = device.get("name", "Unnamed Device")
                logging.debug(f"    Fetching config for device: {device_name} (ID: {device_id})")
                try:
                    # Fetch the device configuration using the Mist API
                    config = mistapi.api.v1.sites.devices.getSiteDevice(apisession, site_id, device_id).data
                    config["site_id"] = site_id
                    config["site_name"] = site_name
                    all_device_configs.append(config)
                    logging.info(f"    ✅ Fetched config for device: {device_name} (ID: {device_id}) at site: {site_name}")
                except Exception as e:
                    logging.warning(f"    ⚠️ Failed to fetch config for device {device_id} at {site_name}: {e}")
        except Exception as e:
            logging.warning(f"  ⚠️ Failed to list devices for site {site_name}: {e}")

    logging.info(f"Fetched configs for {len(all_device_configs)} gateway devices across all sites.")
    return all_device_configs

def export_nac_event_definitions():
    """
    Export NAC (Network Access Control) event definitions to NacEventDefinitions.csv.
    """
    logging.info("Exporting NAC Event Log Definitions...")  # Log start of function
    print("NAC Event Log Definitions:")
    rawdata = mistapi.api.v1.const.nac_events.listNacEventsDefinitions(apisession).data
    # Write the NAC event definitions to a CSV file
    write_data_to_csv(rawdata, "NacEventDefinitions.csv")
    logging.info("✅ NAC event definitions exported to NacEventDefinitions.csv")  # Log completion

def export_client_event_definitions():
    """
    Export client event log definitions to ClientEventDefinitions.csv.
    """
    logging.info("Exporting client event log definitions...")  # Log start of function
    print("Client Event Log Definitions:")
    rawdata = mistapi.api.v1.const.client_events.listClientEventsDefinitions(apisession).data
    write_data_to_csv(rawdata, "ClientEventDefinitions.csv")
    logging.info("✅ Client event definitions exported to ClientEventDefinitions.csv")  # Log completion

def export_device_event_definitions():
    """
    Export device event log definitions to DeviceEventDefinitions.csv.
    """
    logging.info("Exporting device event log definitions...")  # Log start of function
    print("Device Event Log Definitions:")
    rawdata = mistapi.api.v1.const.device_events.listDeviceEventsDefinitions(apisession).data
    # Write the device event definitions to a CSV file
    write_data_to_csv(rawdata, "DeviceEventDefinitions.csv")
    logging.info("✅ Device event definitions exported to DeviceEventDefinitions.csv")  # Log completion

def export_mist_edge_event_definitions():
    """
    Export Mist Edge event log definitions to MistEdgeEventDefinitions.csv.
    """
    logging.info("Exporting Mist Edge event log definitions...")  # Log start of function
    print("Mist Edge Event Log Definitions:")
    rawdata = mistapi.api.v1.const.mxedge_events.listMxEdgeEventsDefinitions(apisession).data
    write_data_to_csv(rawdata, "MistEdgeEventDefinitions.csv")
    logging.info("✅ Mist Edge event definitions exported to MistEdgeEventDefinitions.csv")  # Log completion

def export_other_device_event_definitions():
    """
    Export other device event log definitions to OtherEventDefinitions.csv.
    """
    logging.info("Exporting other device event log definitions...")  # Log start of function
    print("Other Event Log Definitions:")
    # Fetch the other device event definitions using the Mist API
    rawdata = mistapi.api.v1.const.otherdevice_events.listOtherDeviceEventsDefinitions(apisession).data
    # Write the event definitions to a CSV file
    write_data_to_csv(rawdata, "OtherEventDefinitions.csv")
    logging.info("✅ Other device event definitions exported to OtherEventDefinitions.csv")  # Log completion

def export_system_event_definitions():
    """
    Export system event log definitions to SystemEventDefinitions.csv.
    """
    logging.info("Exporting system event log definitions...")  # Log start of function
    print("System Event Log Definitions:")
    rawdata = mistapi.api.v1.const.system_events.listSystemEventsDefinitions(apisession).data
    # Write the system event definitions to a CSV file
    write_data_to_csv(rawdata, "SystemEventDefinitions.csv")
    logging.info("✅ System event definitions exported to SystemEventDefinitions.csv")  # Log completion

def export_alarm_definitions():
    """
    Export alarm log definitions to AlarmDefinitions.csv and display in a PrettyTable.
    Adds logging for each step.
    """
    logging.info("Exporting alarm log definitions...")  # Log start of function
    print("Alarm Log Definitions:")
    # Fetch alarm definitions from the Mist API
    rawdata = mistapi.api.v1.const.alarm_defs.listAlarmDefinitions(apisession).data
    logging.info(f"Fetched {len(rawdata)} alarm definitions from API.")
    # Sort alarm definitions by 'key'
    alarm_defs = sorted(rawdata, key=lambda x: x.get("key", ""))
    # Write alarm definitions to CSV
    write_data_to_csv(alarm_defs, "AlarmDefinitions.csv")
    logging.info("Alarm definitions written to AlarmDefinitions.csv")
    # Prepare PrettyTable for display
    table = PrettyTable()
    table.field_names = ["Key", "Display", "Group", "Severity", "Fields"]
    for alarm in alarm_defs:
        # Add each alarm definition as a row in the table
        table.add_row([
            alarm.get("key"),
            alarm.get("display"),
            alarm.get("group"),
            alarm.get("severity"),
            ", ".join(alarm.get("fields", [])) if isinstance(alarm.get("fields"), list) else alarm.get("fields")
        ])
    logging.info("\n" + table.get_string())  # Log the table output

def export_all_gateway_synthetic_tests():
    """
    Collects and exports synthetic test stats for all gateways in the organization.
    Iterates through all sites with gateways, fetches synthetic test stats for each gateway device,
    and writes the results to AllGatewaySyntheticTests.csv.
    """
    logging.info("[INFO] Collecting synthetic test stats for all gateways in the org...")
    org_id = get_org_id()
    site_ids = get_sites_with_gateways(apisession, org_id)
    all_stats = []

    if not site_ids:
        logging.warning("[WARN] No sites with gateways found. Exiting export_all_gateway_synthetic_tests.")
        return

    for site_id in tqdm(site_ids, desc="Sites", unit="site"):
        try:
            # Fetch all gateway devices for the current site
            response = mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id, type="gateway")
            devices = mistapi.get_all(response=response, mist_session=apisession)
            logging.info(f"[INFO] Found {len(devices)} gateway devices at site {site_id}.")
            for device in tqdm(devices, desc=f"Site {site_id}", unit="device", leave=False):
                device_id = device.get("id")
                device_name = device.get("name", "")
                try:
                    # Fetch synthetic test stats for the gateway device
                    stats = mistapi.api.v1.sites.devices.getSiteDeviceSyntheticTest(apisession, site_id, device_id).data
                    stats["site_id"] = site_id
                    stats["site_name"] = device.get("site_name", "")
                    stats["device_id"] = device_id
                    stats["device_name"] = device_name
                    all_stats.append(stats)
                    logging.info(f"[INFO] Collected synthetic test stats for device {device_name} ({device_id}) at site {site_id}.")
                except Exception as e:
                    logging.warning(f"⚠️ Failed to fetch test stats for device {device_id} at site {site_id}: {e}")
        except Exception as e:
            logging.warning(f"⚠️ Failed to list devices for site {site_id}: {e}")

    if all_stats:
        filename = "AllGatewaySyntheticTests.csv"
        # Flatten nested fields for CSV compatibility
        flattened = flatten_all_nested_fields(all_stats)
        # Escape multiline strings for CSV compatibility
        sanitized = escape_multiline_strings(flattened)
        # Write the processed data to a CSV file
        write_data_to_csv(sanitized, filename)
        logging.info(f"✅ Synthetic test results saved to {filename} ({len(all_stats)} records).")
    else:
        logging.warning("⚠️ No synthetic test results found. CSV not created.")

def get_sites_with_gateways(apisession, org_id):
    """
    Fetches all sites in the organization that have at least one gateway device.

    Args:
        apisession: The Mist API session object.
        org_id: The organization ID.

    Returns:
        List of site IDs that have at least one gateway device.
    """
    logging.info("[INFO] Fetching org inventory to find sites with gateways...")
    # Fetch the full org inventory (all devices)
    response = mistapi.api.v1.orgs.inventory.getOrgInventory(apisession, org_id, limit=1000)
    devices = mistapi.get_all(response=response, mist_session=apisession)
    logging.info(f"[INFO] Retrieved {len(devices)} devices from org inventory.")

    # Collect unique site_ids for devices of type 'gateway'
    gateway_sites = {device["site_id"] for device in devices if device.get("type") == "gateway" and "site_id" in device}
    logging.info(f"[INFO] Found {len(gateway_sites)} sites with at least one gateway.")

    return list(gateway_sites)

def export_all_gateway_test_results_by_site():
    """
    Export all synthetic test results (including speed tests) for all sites with gateways.
    Fetches test results for each site with at least one gateway device and writes them to a CSV.
    """
    logging.info("[INFO] Searching all test results (including speed tests) for sites with gateways...")
    org_id = get_org_id()
    site_ids = get_sites_with_gateways(apisession, org_id)
    all_results = []

    if not site_ids:
        logging.warning("⚠️ No sites with gateways found.")
        return

    for site_id in tqdm(site_ids, desc="Sites", unit="site"):
        try:
            # Fetch synthetic test results for the current site
            response = mistapi.api.v1.sites.synthetic_test.searchSiteSyntheticTest(
                apisession, site_id
            )
            if not hasattr(response, "data"):
                logging.warning(f"⚠️ No data attribute in response for site {site_id}")
                continue

            # Extract results from the response
            results = response.data.get("results", []) if isinstance(response.data, dict) else []
            logging.info(f"[{site_id}] Retrieved {len(results)} test results.")

            for result in results:
                result["site_id"] = site_id  # Annotate result with site_id
                all_results.append(result)

        except Exception as e:
            logging.warning(f"⚠️ Failed to fetch test results for site {site_id}: {e}")

    if all_results:
        filename = "AllGatewayTestResults.csv"
        # Flatten nested fields for CSV compatibility
        flattened = flatten_all_nested_fields(all_results)
        # Escape multiline strings for CSV compatibility
        sanitized = escape_multiline_strings(flattened)
        # Write the processed data to a CSV file
        write_data_to_csv(sanitized, filename)
        logging.info(f"✅ All test results saved to {filename} ({len(all_results)} records).")
    else:
        logging.warning("⚠️ No test results found. CSV not created.")

def export_sites_with_location_info():
    """
    Export a list of sites with location and timezone information to SitesWithLocations.csv.
    Logs each step for traceability.
    """
    logging.info("Listing Sites with Locations:")
    org_id = get_org_id()
    logging.debug(f"Using org_id: {org_id} for site location export.")

    # Fetch all sites in the organization
    response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id)
    sites = mistapi.get_all(response=response, mist_session=apisession)
    logging.info(f"Fetched {len(sites)} sites from the organization.")

    site_data = []
    for site in sites:
        # Extract relevant site information, handling missing fields gracefully
        site_info = {
            "name": site.get("name", ""),
            "address": site.get("address", ""),
            "latitude": site.get("latlng", {}).get("lat", ""),
            "longitude": site.get("latlng", {}).get("lng", ""),
            "timezone": site.get("timezone", "")
        }
        site_data.append(site_info)
        logging.debug(f"Processed site: {site_info['name']} (Lat: {site_info['latitude']}, Lng: {site_info['longitude']})")

    # Escape multiline strings for CSV compatibility
    site_data = escape_multiline_strings(site_data)
    # Write the site data to a CSV file
    write_data_to_csv(site_data, "SitesWithLocations.csv")
    logging.info("Site location data written to SitesWithLocations.csv")

    # Prepare and display a PrettyTable for visual inspection
    table = PrettyTable()
    table.field_names = ["Name", "Address", "Latitude", "Longitude", "Timezone"]
    for site in site_data:
        table.add_row([
            site["name"],
            site["address"],
            site["latitude"],
            site["longitude"],
            site["timezone"]
        ])
    logging.info("\n" + table.get_string())  # Log the table output for reference

def export_gateways_with_site_info():
    """
    Fetches all gateway devices in the organization, enriches them with site and address info,
    and exports the result to GatewaysWithSiteInfo.csv. Also logs and displays a summary table.
    """
    logging.info("Fetching Gateways with Site Info...")
    org_id = get_org_id()

    # Fetch site list and build a lookup dictionary for site info
    site_response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id)
    sites = mistapi.get_all(response=site_response, mist_session=apisession)
    site_lookup = {
        site["id"]: {
            "name": site.get("name", ""),
            "address": site.get("address", "")
        } for site in sites
    }
    logging.debug(f"Loaded {len(site_lookup)} sites for lookup.")

    # Fetch org inventory (all devices)
    inv_response = mistapi.api.v1.orgs.inventory.getOrgInventory(apisession, org_id)
    inventory = mistapi.get_all(response=inv_response, mist_session=apisession)
    logging.debug(f"Loaded {len(inventory)} devices from org inventory.")

    def split_address(address):
        """
        Splits a full address string into street, city, state, zip, and country.
        Returns empty strings if parsing fails.
        """
        try:
            parts = address.split(", ")
            street = parts[0]
            city = parts[1]
            state_zip = parts[2].split()
            state = state_zip[0]
            zip_code = state_zip[1]
            country = parts[3]
            return street, city, state, zip_code, country
        except Exception as e:
            logging.debug(f"Failed to split address '{address}': {e}")
            return address, "", "", "", ""

    # Filter for gateways and enrich with site info
    gateways = []
    for device in tqdm(inventory, desc="Processing Gateways", unit="device"):
        if device.get("type") == "gateway":
            site_id = device.get("site_id")
            site_info = site_lookup.get(site_id, {"name": "Unknown", "address": "Unknown"})
            device["site_name"] = site_info["name"]
            device["site_address"] = site_info["address"]
            street, city, state, zip_code, country = split_address(site_info["address"])
            device["street"] = street
            device["city"] = city
            device["state"] = state
            device["zip_code"] = zip_code
            device["country"] = country
            gateways.append(device)
    logging.info(f"Enriched {len(gateways)} gateway devices with site info.")

    # Flatten nested fields and escape multiline strings for CSV compatibility
    gateways = flatten_all_nested_fields(gateways)
    gateways = escape_multiline_strings(gateways)
    gateways = sorted(gateways, key=lambda x: x.get("site_name", ""))
    write_data_to_csv(gateways, "GatewaysWithSiteInfo.csv")
    logging.info("Gateway data written to GatewaysWithSiteInfo.csv")

    # Display a summary table in logs
    table = PrettyTable()
    table.field_names = ["name", "mac", "model", "serial", "site_name", "street", "city", "state", "zip_code", "country"]
    for gw in gateways:
        table.add_row([
            gw.get("name", ""),
            gw.get("mac", ""),
            gw.get("model", ""),
            gw.get("serial", ""),
            gw.get("site_name", ""),
            gw.get("street", ""),
            gw.get("city", ""),
            gw.get("state", ""),
            gw.get("zip_code", ""),
            gw.get("country", "")
        ])
    logging.info("\n" + table.get_string())

def export_all_devices_with_site_info():
    """
    Fetches all devices in the organization, enriches them with site and address info,
    and exports the result to AllDevicesWithSiteInfo.csv. Also logs and displays a summary table.
    """
    logging.info("Fetching All Devices with Site Info...")  # Log start of function
    org_id = get_org_id()

    # Fetch all sites and build a lookup dictionary for site info
    site_response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id)
    sites = mistapi.get_all(response=site_response, mist_session=apisession)
    site_lookup = {
        site["id"]: {
            "name": site.get("name", ""),
            "address": site.get("address", "")
        } for site in sites
    }
    logging.debug(f"Loaded {len(site_lookup)} sites for lookup.")

    # Fetch org inventory (all devices)
    inv_response = mistapi.api.v1.orgs.inventory.getOrgInventory(apisession, org_id)
    inventory = mistapi.get_all(response=inv_response, mist_session=apisession)
    logging.debug(f"Loaded {len(inventory)} devices from org inventory.")

    def split_address(address):
        """
        Splits a full address string into street, city, state, zip, and country.
        Returns empty strings if parsing fails.
        """
        try:
            parts = address.split(", ")
            street = parts[0]
            city = parts[1]
            state_zip = parts[2].split()
            state = state_zip[0]
            zip_code = state_zip[1]
            country = parts[3]
            return street, city, state, zip_code, country
        except Exception as e:
            logging.debug(f"Failed to split address '{address}': {e}")
            return address, "", "", "", ""

    enriched_devices = []
    for device in tqdm(inventory, desc="Processing Devices", unit="device"):
        site_id = device.get("site_id")
        site_info = site_lookup.get(site_id, {"name": "Unknown", "address": "Unknown"})
        device["site_name"] = site_info["name"]
        device["site_address"] = site_info["address"]
        street, city, state, zip_code, country = split_address(site_info["address"])
        device["street"] = street
        device["city"] = city
        device["state"] = state
        device["zip_code"] = zip_code
        device["country"] = country
        enriched_devices.append(device)
        logging.debug(f"Enriched device {device.get('name', '')} ({device.get('mac', '')}) with site info.")

    # Flatten nested fields and escape multiline strings for CSV compatibility
    enriched_devices = flatten_all_nested_fields(enriched_devices)
    enriched_devices = escape_multiline_strings(enriched_devices)
    enriched_devices = sorted(enriched_devices, key=lambda x: x.get("site_name", ""))
    write_data_to_csv(enriched_devices, "AllDevicesWithSiteInfo.csv")
    logging.info(f"All device data written to AllDevicesWithSiteInfo.csv ({len(enriched_devices)} records).")

    # Display a summary table in logs
    table = PrettyTable()
    table.field_names = ["name", "mac", "model", "serial", "type", "site_name", "street", "city", "state", "zip_code", "country"]
    for dev in enriched_devices:
        table.add_row([
            dev.get("name", ""),
            dev.get("mac", ""),
            dev.get("model", ""),
            dev.get("serial", ""),
            dev.get("type", ""),
            dev.get("site_name", ""),
            dev.get("street", ""),
            dev.get("city", ""),
            dev.get("state", ""),
            dev.get("zip_code", ""),
            dev.get("country", "")
        ])
    logging.info("\n" + table.get_string())  # Log the table output for reference

def generate_support_package():
    logging.info("Generating support package for each site...")

    # List of required CSV files and their generation functions
    required_files = [
        ("OrgAlarms.csv", search_org_alarms),
        ("OrgDeviceEvents.csv", export_recent_device_events),
        ("SiteList.csv", export_org_site_list),
        ("OrgDevices.csv", export_all_org_devices),
        ("OrgDeviceStats.csv", export_org_device_statistics),
        ("OrgDevicePortStats.csv", export_org_device_port_stats),
        ("AllGatewayTestResults.csv", export_all_gateway_test_results_by_site),
    ]

    # Ensure all required files are fresh or regenerate them
    for filename, func in required_files:
        logging.debug(f"Checking freshness of {filename}...")
        check_and_generate_csv(filename, func, freshness_minutes=15)

    # Ensure SiteList.csv is generated before loading
    check_and_generate_csv('SiteList.csv', export_org_site_list, freshness_minutes=15)

    # Load the pulled data into dictionaries
    logging.debug("Loading CSV data into dictionaries for support package assembly...")
    site_data = load_csv_into_dict('SiteList.csv', 'id')
    alarms_data = load_csv_into_dict('OrgAlarms.csv', 'site_id')
    events_data = load_csv_into_dict('OrgDeviceEvents.csv', 'site_id')
    devices_data = load_csv_into_dict('OrgDevices.csv', 'name')
    device_stats_data = load_csv_into_dict('OrgDeviceStats.csv', 'site_id')
    port_stats_data = load_csv_into_dict('OrgDevicePortStats.csv', 'site_id')

    # Load speedtest data if available
    if os.path.exists('AllGatewayTestResults.csv'):
        logging.debug("Loading AllGatewayTestResults.csv for speedtest data...")
        speedtest_data = load_csv_into_dict('AllGatewayTestResults.csv', 'site_id')
    else:
        logging.warning("⚠️ AllGatewayTestResults.csv not found. Skipping speedtest data.")
        speedtest_data = {}

    # Create a support package for each site with alarms or events
    for site_id, site_info in site_data.items():
        # Only generate support package if there are alarms or events for the site
        if not alarms_data.get(site_id) and not events_data.get(site_id):
            logging.info(f"Skipping site {site_id} — no alarms or events.")
            continue

        logging.info(f"Generating support package for site: {site_id}")
        # Gather all relevant data for the site
        support_data = {
            'alarms': alarms_data.get(site_id, []),
            'events': events_data.get(site_id, []),
            'devices': devices_data.get(site_id, []),
            'device_stats': device_stats_data.get(site_id, []),
            'port_stats': port_stats_data.get(site_id, []),
            'speedtests': speedtest_data.get(site_id, []),
        }

        support_package_filename = f"SupportPackage_{site_id}.csv"
        logging.debug(f"Writing support package to {support_package_filename}...")
        write_support_package_to_csv(support_data, support_package_filename)
        logging.info(f"Support package written for site {site_id}.")

    logging.info("✅ Support packages generated for applicable sites.")
    logging.info("✅ Support packages generated for all sites!")

def load_csv_into_dict(filename, key):
    """
    Loads CSV data into a dictionary keyed by the specified column.
    Each key maps to a list of rows (as dictionaries) that share the same key value.
    Adds logging for file loading and key distribution.
    """
    logging.info(f"Loading CSV file '{filename}' into dictionary keyed by '{key}'...")
    with open(filename, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)  # Create a CSV reader
        data_dict = {}  # Initialize an empty dictionary
        row_count = 0
        for row in reader:
            data_key = row.get(key)  # Get the value to use as the key
            if data_key is None:
                logging.warning(f"Row missing key '{key}': {row}")
                continue
            if data_key not in data_dict:
                data_dict[data_key] = []  # Initialize a list for this key
            data_dict[data_key].append(row)  # Add the row to the dictionary
            row_count += 1
        logging.info(f"Loaded {row_count} rows from '{filename}'. Found {len(data_dict)} unique keys for '{key}'.")
    return data_dict  # Return the dictionary

def write_support_package_to_csv(data, filename):
    """
    Writes the support package data (a dict of lists of dicts) to a CSV file.
    Each section in 'data' is a list of dictionaries. All unique keys across all sections are used as CSV columns.
    """
    logging.debug(f"Preparing to write support package to {filename}...")

    fieldnames = set()  # Initialize a set to collect all field names
    # Collect all unique field names from all sections
    for section_name, section in data.items():
        logging.debug(f"Processing section '{section_name}' with {len(section)} rows.")
        for row in section:
            fieldnames.update(row.keys())  # Add all keys to the fieldnames set
    fieldnames = sorted(fieldnames)  # Sort the fieldnames for consistent column order

    logging.debug(f"Final CSV fieldnames: {fieldnames}")

    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)  # Create a CSV writer
        writer.writeheader()  # Write the header row
        row_count = 0
        for section_name, section in data.items():
            for row in section:
                writer.writerow(row)  # Write each row to the CSV file
                row_count += 1
        logging.info(f"Wrote {row_count} rows to {filename} for support package.")

    logging.info(f"Support package written to {filename}")  # Log completion of the file write

def poll_marvis_actions():
    """
    Polls Marvis actions for the organization, filters for open actions, and writes them to a CSV.
    Adds logging and comments for traceability.
    """
    logging.info("🔍 Polling Marvis Actions...")
    print("🔍 Polling Marvis Actions...")
    org_id = get_org_id()
    logging.debug(f"Using org_id: {org_id} for Marvis actions polling.")

    # Call the Mist API to get Marvis actions
    response = mistapi.api.v1.orgs.troubleshoot.troubleshootOrg(apisession, org_id)
    rawdata = mistapi.get_all(response=response, mist_session=apisession)
    logging.info(f"Fetched {len(rawdata)} Marvis actions from API.")

    # Filter only open actions (state == "open")
    open_actions = [action for action in rawdata if action.get("state") == "open"]
    logging.info(f"Filtered {len(open_actions)} open Marvis actions.")

    # Flatten and clean the data for CSV compatibility
    data = flatten_all_nested_fields(open_actions)
    data = escape_multiline_strings(data)
    logging.debug("Flattened and sanitized open Marvis actions for CSV.")

    # Write to CSV
    write_data_to_csv(data, "OpenMarvisActions.csv")
    logging.info(f"✅ {len(open_actions)} open Marvis actions written to OpenMarvisActions.csv")
    print(f"✅ {len(open_actions)} open Marvis actions written to OpenMarvisActions.csv")

def export_current_guests():
    """
    Export all current guest users in the org to OrgCurrentGuests.csv
    """
    logging.info("Exporting all current guest users in the org...")  # Log start of function
    org_id = get_org_id()
    logging.debug(f"Using org_id: {org_id} for current guest export.")

    # Call the Mist API to get current guest authorizations
    response = mistapi.api.v1.orgs.guests.searchOrgGuestAuthorization(apisession, org_id, limit=1000)
    guests = mistapi.get_all(response=response, mist_session=apisession)
    logging.info(f"Fetched {len(guests)} current guest users from API.")

    # Flatten nested fields for CSV compatibility
    guests = flatten_all_nested_fields(guests)
    # Escape multiline strings for CSV compatibility
    guests = escape_multiline_strings(guests)

    # Write the processed data to a CSV file
    write_data_to_csv(guests, "OrgCurrentGuests.csv")
    logging.info("✅ Current guests exported to OrgCurrentGuests.csv")  # Log completion

def export_historical_guests():
    """
    Export all guest users from the last 7 days to OrgHistoricalGuests.csv
    """
    logging.info("Exporting all guest users from the last 7 days...")  # Log start of function
    org_id = get_org_id()
    # Calculate epoch for 7 days ago
    end_time = int(time.time())
    start_time = end_time - 7 * 24 * 3600
    logging.debug(f"Fetching guest authorizations from {start_time} to {end_time} (epoch seconds).")
    # Call the Mist API to get guest authorizations in the last 7 days
    response = mistapi.api.v1.orgs.guests.searchOrgGuestAuthorization(
        apisession, org_id, limit=1000, start=start_time, end=end_time
    )
    guests = mistapi.get_all(response=response, mist_session=apisession)
    logging.info(f"Fetched {len(guests)} historical guest users from API.")
    # Flatten nested fields for CSV compatibility
    guests = flatten_all_nested_fields(guests)
    # Escape multiline strings for CSV compatibility
    guests = escape_multiline_strings(guests)
    # Write the processed data to a CSV file
    write_data_to_csv(guests, "OrgHistoricalGuests.csv")
    logging.info("✅ Historical guests exported to OrgHistoricalGuests.csv")  # Log completion

def export_all_switch_vc_stats():
    """
    Export virtual chassis stats (including stacking cable info) for all switches in the org.
    """
    logging.info("Exporting all switch virtual chassis stats...")

    # Ensure OrgInventory.csv is fresh
    check_and_generate_csv("OrgInventory.csv", export_org_device_inventory, freshness_minutes=15)

    # Load OrgInventory.csv and filter for switches
    with open("OrgInventory.csv", mode="r", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        switches = [row for row in reader if row.get("type") == "switch"]

    if not switches:
        logging.warning("No switches found in OrgInventory.csv.")
        return

    all_vc_stats = []

    for switch in tqdm(switches, desc="Switches", unit="switch"):
        site_id = switch.get("site_id")
        device_id = switch.get("id")
        name = switch.get("name", "")
        mac = switch.get("mac", "")
        model = switch.get("model", "")
        serial = switch.get("serial", "")

        # Log which switch is being processed
        logging.debug(f"Processing switch: name={name}, id={device_id}, site_id={site_id}, mac={mac}, model={model}, serial={serial}")

        if not site_id or not device_id:
            logging.warning(f"Skipping switch with missing site_id or device_id: name={name}, mac={mac}")
            continue

        try:
            # Get VC stats for this switch (returns a flat dict)
            vc_stats = mistapi.api.v1.sites.devices.getSiteDeviceVirtualChassis(apisession, site_id, device_id).data
            logging.debug(f"Fetched VC stats for switch {name} ({device_id}): {vc_stats}")
            # Merge all switch info and VC info into a single dictionary
            entry = {**switch, **vc_stats}
            all_vc_stats.append(entry)
        except Exception as e:
            logging.warning(f"Failed to fetch VC stats for switch {name} ({device_id}): {e}")

    # Flatten and write to CSV
    logging.info(f"Flattening and sanitizing {len(all_vc_stats)} VC stats entries for CSV export.")
    all_vc_stats = flatten_all_nested_fields(all_vc_stats)
    all_vc_stats = escape_multiline_strings(all_vc_stats)
    write_data_to_csv(all_vc_stats, "OrgSwitchVCStats.csv")
    logging.info(f"✅ Switch VC stats exported to OrgSwitchVCStats.csv ({len(all_vc_stats)} records).")
    # Optionally log a preview of the data
    if all_vc_stats:
        logging.debug(f"Sample VC stats row: {all_vc_stats[0]}")
        # Display a summary PrettyTable for quick inspection
        table = PrettyTable()
        summary_fields = ["name", "mac", "model", "serial", "site_id", "vc_mac", "status", "members_0_vc_role", "members_1_vc_role"]
        table.field_names = [f for f in summary_fields if f in all_vc_stats[0]]
        for row in all_vc_stats:
            table.add_row([row.get(f, "") for f in table.field_names])
        logging.info("\n" + table.get_string())

def select_site_and_device(site_id=None, device_id=None):
    """
    Returns site_id and device_id, either from arguments or via interactive prompts.
    """
    if not site_id:
        site_id = prompt_user_to_select_site_id_from_csv()
        if not site_id:
            print("❌ No site selected.")
            return None, None

    if not device_id:
        device_id = prompt_user_to_select_device_id(site_id, device_type=device_type)
        if not device_id:
            print("❌ No device selected.")
            return None, None

    return site_id, device_id

def create_shell_session(site_id, device_id):
    """
    Creates a shell session and returns the WebSocket URL.
    """
    try:
        resp = mistapi.api.v1.sites.devices.createSiteDeviceShellSession(apisession, site_id, device_id)
        shell_data = resp.data
        return shell_data.get("url")
    except Exception as e:
        print(f"❌ Failed to create shell session: {e}")
        return None

def run_interactive_shell(shell_url, debug=False):
    import json, sys, shutil, threading, time
    from sshkeyboard import listen_keyboard, stop_listening
    import websocket
    import pyte

    if debug:
        websocket.enableTrace(True)

    print("🔌 Connecting to WebSocket shell...")
    ws = websocket.create_connection(shell_url)
    print("🟢 Connected.")

    screen = pyte.Screen(80, 40)
    stream = pyte.Stream(screen)

    def _resize():
        cols, rows = shutil.get_terminal_size()
        resize_msg = json.dumps({'resize': {'width': cols, 'height': rows}})
        if debug:
            print(f"[DEBUG] Sending resize: {resize_msg}")
        ws.send(resize_msg)

    def _ws_in():
        while ws.connected:
            try:
                data = ws.recv()
                if isinstance(data, bytes):
                    data = data.decode('utf-8', errors='ignore')
                if debug:
                    print(f"[DEBUG] Raw recv: {repr(data)}")
                if data:
                    stream.feed(data)
                    for y in sorted(screen.dirty):
                        sys.stdout.write(f"\x1b[{y+1};1H")  # Move cursor to line y+1
                        sys.stdout.write(screen.display[y] + "\x1b[K")  # Clear to end of line
                    sys.stdout.flush()
                    screen.dirty.clear()
            except Exception as e:
                print(f'\n## Connection lost: {e} ##')
                return
    def _ws_out(key):
        if ws.connected:
            keymap = {
                "enter": "\n", "space": " ", "tab": "\t",
                "up": "\x00\x1b[A", "down": "\x00\x1b[B",
                "left": "\x00\x1b[D", "right": "\x00\x1b[C",
                "backspace": "\x08"
            }
            if key == "~":
                print('\n## Exit from shell ##')
                ws.sock.shutdown(2)
                ws.sock.close()
                stop_listening()
                return
            k = keymap.get(key, key)
            data = f"\00{k}"
            data_byte = bytearray(map(ord, data))
            if debug:
                print(f"[DEBUG] Sending: {repr(data)}")
            try:
                ws.send_binary(data_byte)
            except Exception as e:
                print(f'\n## Send failed: {e} ##')
                return

    _resize()
    threading.Thread(target=_ws_in).start()

    # Wake up Juniper SSR prompt
    time.sleep(1)
    ws.send_binary(bytearray(map(ord, "\00\n\n")))
    if debug:
        print("[DEBUG] Sent wakeup sequence to Juniper SSRs")

    listen_keyboard(on_release=_ws_out, delay_second_char=0, delay_other_chars=0, lower=False)


    _resize()
    threading.Thread(target=_ws_in).start()

    # Wake up Juniper SSR prompt
    time.sleep(1)
    ws.send_binary(bytearray(map(ord, "\00\n\n")))
    if debug:
        print("[DEBUG] Sent wakeup sequence to Juniper SSRs")

    listen_keyboard(on_release=_ws_out, delay_second_char=0, delay_other_chars=0, lower=False)

def launch_cli_shell(site_id=None, device_id=None, debug=False):
    site_id, device_id = select_site_and_device(site_id, device_id)
    if not site_id or not device_id:
        return
    shell_url = create_shell_session(site_id, device_id)
    if shell_url:
        run_interactive_shell(shell_url, debug=debug)

def listen_for_command_output(mist_host, mist_apitoken, site_id, device_id, session_id, timeout=30, idle_timeout=3):
    import websocket
    import json
    import threading
    import time
    import logging
    from prettytable import PrettyTable

    ws_url = f"wss://{mist_host}/api-ws/v1/stream"
    headers = [f"Authorization: Token {mist_apitoken}"]
    subscribe_msg = {
        "subscribe": f"/sites/{site_id}/devices/{device_id}/cmd"
    }

    output_lines = []
    buffer = ""
    last_message_time = time.time()

    def on_message(ws, message):
        nonlocal last_message_time, buffer, output_lines
        try:
            last_message_time = time.time()
            logging.info(f"🔔 Raw WebSocket message:\n{message}")

            msg = json.loads(message)
            data_str = msg.get("data", "{}")
            data_obj = json.loads(data_str) if isinstance(data_str, str) else data_str
            inner_data = data_obj.get("data", {})
            if isinstance(inner_data, str):
                inner_data = json.loads(inner_data)

            if inner_data.get("session") == session_id:
                raw_output = inner_data.get("raw", "")
                buffer += raw_output

                while "\n" in buffer:
                    line, buffer = buffer.split("\n", 1)
                    output_lines.append(line)

        except Exception as e:
            logging.warning(f"⚠️ Error parsing message: {e}")

    def on_error(ws, error):
        logging.error(f"❌ WebSocket error: {error}")

    def on_close(ws, *args):
        logging.info("🔌 WebSocket closed.")
        if output_lines:
            compiled_output = "\n".join(output_lines)
            print("\n📥 ARP Output Received:\n")
            rows = compiled_output.split("\n")
            parsed_rows = [row.split("\t") for row in rows if row.strip()]
            max_cols = max(len(row) for row in parsed_rows)
            for row in parsed_rows:
                while len(row) < max_cols:
                    row.append("")
            table = PrettyTable()
            table.field_names = [f"Col {i+1}" for i in range(max_cols)]
            for row in parsed_rows:
                table.add_row(row)
            print(table)
            logging.info(f"📥 Compiled ARP Output:\n{compiled_output}")
        else:
            print("⚠️ No ARP output received for this session.")
            logging.warning("⚠️ No ARP output received for this session.")

    def on_open(ws):
        logging.info("🔓 WebSocket opened. Subscribing...")
        ws.send(json.dumps(subscribe_msg))

    ws = websocket.WebSocketApp(
        ws_url,
        header=headers,
        on_message=on_message,
        on_error=on_error,
        on_close=on_close,
        on_open=on_open
    )

    def run_ws():
        ws.run_forever()

    ws_thread = threading.Thread(target=run_ws)
    ws_thread.start()

    start_time = time.time()
    while time.time() - start_time < timeout:
        time.sleep(1)
        if time.time() - last_message_time > idle_timeout and output_lines:
            logging.info("⏹️ Idle timeout reached. Closing WebSocket.")
            ws.close()
            break

    if ws.keep_running:
        logging.warning("⏱️ Timeout waiting for ARP output.")
        ws.close()



def trigger_arp_command(mist_host, mist_apitoken, site_id, device_id):
    import requests

    url = f"https://{mist_host}/api/v1/sites/{site_id}/devices/{device_id}/arp"
    headers = {'Authorization': f'Token {mist_apitoken}'}
    response = requests.post(url, headers=headers, json={})

    if response.status_code == 200:
        session_id = response.json().get("session")
        print(f"✅ ARP command triggered. Session ID: {session_id}")
        return session_id
    else:
        print(f"❌ Failed to trigger ARP command: {response.status_code}")
        print(response.text)
        return None

def run_arp_via_websocket(site_id=None, device_id=None):
    if not site_id or not device_id:
        site_id, device_id = select_site_and_device(site_id, device_id)
    if not site_id or not device_id:
        return

    # Retrieve mist_host and mist_apitoken from the apisession or environment
    mist_host = getattr(apisession, "host", None) or os.getenv("MIST_HOST")
    mist_apitoken = getattr(apisession, "apitoken", None) or os.getenv("MIST_APITOKEN")

    if not mist_host or not mist_apitoken:
        print("❌ Mist host or API token not found in session or environment.")
        return

    print("🔌 Subscribing to WebSocket stream...")
    session_id = trigger_arp_command(mist_host, mist_apitoken, site_id, device_id)
    if session_id:
        listen_for_command_output(mist_host.replace("api.", "api-ws."), mist_apitoken, site_id, device_id, session_id)


def _load_env(env_file: str, mist_host: str, mist_apitoken: str, mist_site_id: str, mist_device_id: str = ""):
    from dotenv import load_dotenv
    import os

    if env_file.startswith("~/"):
        env_file = os.path.join(os.path.expanduser("~"), env_file.replace("~/", ""))
    load_dotenv(dotenv_path=env_file, override=True)

    mist_host = os.getenv("MIST_HOST", mist_host)
    mist_apitoken = os.getenv("MIST_APITOKEN", mist_apitoken)
    mist_site_id = os.getenv("MIST_SITE_ID", mist_site_id)
    mist_device_id = os.getenv("MIST_DEVICE_ID", mist_device_id)

    return mist_host, mist_apitoken, mist_site_id, mist_device_id



menu_actions = {
    # 🗂️ Setup & Core Logs
    "0": (select_site, "Select a site (used by other functions)"),
    "1": (search_org_alarms, "Export all organization alarms from the past day"),
    "2": (export_recent_device_events, "Export all device events from the past 24 hours"),
    "3": (export_org_audit_logs, "Export audit logs for the organization"),

    # 📚 Event & Alarm Definitions
    "4": (export_nac_event_definitions, "Export NAC (Network Access Control) event definitions"),
    "5": (export_client_event_definitions, "Export client event definitions"),
    "6": (export_device_event_definitions, "Export device event definitions"),
    "7": (export_mist_edge_event_definitions, "Export Mist Edge event definitions"),
    "8": (export_other_device_event_definitions, "Export other device event definitions"),
    "9": (export_system_event_definitions, "Export system event definitions"),
    "10": (export_alarm_definitions, "Export alarm definitions with severity and field info"),

    # 🏢 Organization-Level Exports
    "11": (export_org_site_list, "Export a list of all sites in the organization"),
    "12": (export_org_device_inventory, "Export the full inventory of devices in the organization"),
    "13": (export_org_device_statistics, "Export statistics for all devices in the organization"),
    "14": (export_org_device_port_stats, "Export port-level statistics for switches and gateways"),
    "15": (export_org_vpn_peer_stats, "Export VPN peer path statistics for the organization"),

    # 🧭 Interactive Site/Device Exploration
    "16": (interactive_view_site_inventory, "View device inventory for a selected site"),
    "17": (interactive_view_device_stats, "View statistics for a selected device at a site"),
    "18": (interactive_view_device_tests, "View synthetic test stats for a selected gateway device"),
    "19": (interactive_view_device_config, "View configuration details for a selected device"),

    # 🌐 Gateway & Site-Wide Exports
    "20": (export_all_gateway_synthetic_tests, "Export synthetic test results for all gateways"),
    "21": (export_all_org_devices, "Export a list of all devices in the organization"),
    "22": (export_all_site_settings, "Export configuration settings for all sites"),
    "23": (export_all_gateway_device_configs, "WIP Export configuration details for all gateway devices across all sites"),
    "24": (export_all_gateway_test_results_by_site, "Export all synthetic test results (including speed tests) for gateways"),

    # 🗺️ Location-Enriched Exports
    "25": (export_sites_with_location_info, "Export a list of sites with location and timezone info"),
    "26": (export_gateways_with_site_info, "Export a list of gateways with associated site and address info"),
    "27": (export_all_devices_with_site_info, "Export a list of all devices with associated site and address info"),
    "28": (process_and_merge_csv_for_sfp_address, "Process and merge CSV files of SFP Module locations into a single CSV file"),
    "29": (generate_support_package, "Generate support package for each site"),
    "30": (poll_marvis_actions, "Poll Marvis actions and export open actions to CSV"),
    "31": (lambda: (export_current_guests(), export_historical_guests()),"Export all current guest users and last 7 days of historical guests to CSV"),
    "32": (export_all_switch_vc_stats, "Export all switch virtual chassis (VC/stacking) stats to CSV"),
    "33": (launch_cli_shell, "Interactively execute a CLI command on a gateway or switch (exit with ~)"),
    "34": (run_arp_via_websocket, "Run ARP command on a device and receive output via WebSocket")
}

def main():
    # --- CLI Argument Parsing ---
    parser = argparse.ArgumentParser(description="MistHelper CLI Interface")
    parser.add_argument("-O", "--org", help="Organization ID")
    parser.add_argument("-M", "--menu", help="Menu option number to execute")
    parser.add_argument("-S", "--site", help="Human-readable site name")
    parser.add_argument("-D", "--device", help="Human-readable device name")
    parser.add_argument("-P", "--port", help="Port ID")
    parser.add_argument("--debug", action="store_true", help="Enable debug output")

    args = parser.parse_args()

    global org_id

    # If any CLI args are passed, override interactive mode
    if len(sys.argv) > 1:
        logging.info("CLI arguments detected, running in non-interactive mode.")
        if args.org:
            org_id = args.org
            logging.info(f"Overriding org_id with CLI argument: {org_id}")
        else:
            org_id = get_org_id()

        # Resolve site name to site_id if needed
        if args.site:
            logging.info(f"Resolving site name '{args.site}' to site_id...")
            sites = mistapi.get_all(mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id), apisession)
            site_lookup = {site["name"]: site["id"] for site in sites}
            site_id = site_lookup.get(args.site)
            if not site_id:
                logging.error(f"❌ Site name '{args.site}' not found.")
                print(f"❌ Site name '{args.site}' not found.")
                sys.exit(1)
            else:
                logging.info(f"Resolved site name '{args.site}' to site_id '{site_id}'.")
        else:
            site_id = None

        # Resolve device name to device_id if needed
        if args.device and site_id:
            logging.info(f"Resolving device name '{args.device}' at site_id '{site_id}'...")
            devices = mistapi.get_all(mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id), apisession)
            device_lookup = {dev["name"]: dev["id"] for dev in devices}
            device_id = device_lookup.get(args.device)
            if not device_id:
                logging.error(f"❌ Device name '{args.device}' not found at site '{args.site}'.")
                print(f"❌ Device name '{args.device}' not found at site '{args.site}'.")
                sys.exit(1)
            else:
                logging.info(f"Resolved device name '{args.device}' to device_id '{device_id}'.")
        else:
            device_id = None

        # Execute the selected menu action
        if args.menu in menu_actions:
            func, _ = menu_actions[args.menu]
            logging.info(f"Executing menu action '{args.menu}'.")

            # Dynamically pass only accepted arguments
            import inspect
            func_args = {
                "site_id": site_id,
                "device_id": device_id,
                "port": args.port,
                "org_id": org_id,
                "debug": args.debug
            }
            sig = inspect.signature(func)
            accepted_args = {
                k: v for k, v in func_args.items()
                if k in sig.parameters and v is not None
            }
            func(**accepted_args)
        else:
            logging.error(f"❌ Invalid menu option: {args.menu}")
            print(f"❌ Invalid menu option: {args.menu}")
            sys.exit(1)

        logging.info("CLI execution complete. Exiting.")
        sys.exit(0)

    # --- Interactive Menu Fallback ---
    if len(sys.argv) == 1:
        logging.info("No CLI arguments detected, running in interactive menu mode.")
        print("\nAvailable Options:")
        for key, (func, description) in menu_actions.items():
            print(f"{key}: {description}")
        iwant = input("\nEnter your selection number now: ").strip()
        selected = menu_actions.get(iwant)
        if selected:
            func, _ = selected
            logging.info(f"User selected menu option '{iwant}'. Executing associated function.")
            func()
            sys.exit(0)
        else:
            logging.warning(f"Invalid selection '{iwant}' entered by user.")
            print("Invalid selection. Please try again.")
            sys.exit(1)


if __name__ == "__main__":
    main()