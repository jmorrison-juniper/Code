import mistapi, csv, ast, json, time, logging, os, argparse, sys
from prettytable import PrettyTable
from tqdm import tqdm
from datetime import datetime, timedelta

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
    from datetime import datetime, timedelta
    import os

    if os.path.exists(file_name):
        file_mtime = datetime.fromtimestamp(os.path.getmtime(file_name))
        if datetime.now() - file_mtime < timedelta(minutes=freshness_minutes):
            logging.info(f"✅ Using cached {file_name} (fresh)")
            return
        else:
            logging.info(f"♻️ {file_name} is older than {freshness_minutes} minutes. Regenerating...")
    else:
        logging.info(f"📄 {file_name} not found. Generating...")

    generate_function()

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
     # Automatically generate missing files
    if not os.path.exists('OrgDevicePortStats.csv'):
        print("⚠️ OrgDevicePortStats.csv not found. Generating it now...")
        export_org_device_port_stats()

    if not os.path.exists('AllDevicesWithSiteInfo.csv'):
        print("⚠️ AllDevicesWithSiteInfo.csv not found. Generating it now...")
        export_all_devices_with_site_info()


    # Load site and device info
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
    global org_id
    if org_id:
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
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_nested_dict(v, new_key, sep=sep).items())
        elif isinstance(v, list):
            if all(isinstance(i, dict) for i in v):
                for idx, item in enumerate(v):
                    items.extend(flatten_nested_dict(item, f"{new_key}{sep}{idx}", sep=sep).items())
            else:
                items.append((new_key, ','.join(map(str, v))))
        else:
            items.append((new_key, v))
    return dict(items)

def flatten_all_nested_fields(data):
    flattened = []
    for entry in data:
        new_entry = {}
        for key, value in entry.items():
            # Try to parse stringified dicts/lists
            if isinstance(value, str) and (value.startswith("{") or value.startswith("[")):
                try:
                    value = ast.literal_eval(value)
                except Exception:
                    try:
                        value = json.loads(value)
                    except Exception:
                        pass  # Leave as string if parsing fails

            # Flatten if it's a dict or list of dicts
            if isinstance(value, dict):
                flat = flatten_nested_dict(value, parent_key=key)
                new_entry.update(flat)
            elif isinstance(value, list):
                if all(isinstance(i, dict) for i in value):
                    for idx, item in enumerate(value):
                        flat = flatten_nested_dict(item, parent_key=f"{key}_{idx}")
                        new_entry.update(flat)
                else:
                    new_entry[key] = ','.join(map(str, value))
            else:
                new_entry[key] = value
        flattened.append(new_entry)
    return flattened

def convert_list_values_to_strings(data):
    for entry in data:
        for key, value in entry.items():
            if isinstance(value, list):
                entry[key] = ','.join(map(str, value))
    return data

def get_all_unique_keys(data):
    fields = set()
    for entry in data:
        fields.update(entry.keys())
    return sorted(fields)

def escape_multiline_strings(data):
    for entry in data:
        for key, value in entry.items():
            if isinstance(value, list):
                entry[key] = ','.join(map(str, value))
            elif isinstance(value, str):
                entry[key] = value.replace('\n', '\\n').replace('\r', '')
    return data

def write_data_to_csv(data, csv_file):
    data = escape_multiline_strings(data)
    fields = get_all_unique_keys(data)
    with open(csv_file, 'w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        for row in data:
            writer.writerow({field: row.get(field, "") for field in fields})
    logging.info(f"Data saved to {csv_file}")

def fetch_process_and_display_data(title, api_call, filename, sort_key=None, display_fields=None, **kwargs):
    print(title)
    org_id = get_org_id()
    response = api_call(apisession, org_id, **kwargs)
    rawdata = mistapi.get_all(response=response, mist_session=apisession)
    data = [entry for entry in rawdata if isinstance(entry, dict)]

    if sort_key:
        data = sorted(data, key=lambda x: x.get(sort_key, ""))

    data = flatten_all_nested_fields(data)
    data = escape_multiline_strings(data)
    fields = get_all_unique_keys(data)
    write_data_to_csv(data, filename)

    table = PrettyTable()
    table.field_names = display_fields if display_fields else fields
    table.valign = "t"
    for item in tqdm(data, desc="Processing", unit="record"):
        row = [item.get(field, "") for field in table.field_names]
        table.add_row(row)
    logging.info("\n" + table.get_string())

def prompt_user_to_select_device_id(site_id, device_type="all", csv_filename="SiteInventory.csv"):
    rawdata = mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id, type=device_type).data
    if not rawdata:
        print("No devices found for the selected site.")
        return None

    inventory = sorted(rawdata, key=lambda x: x.get("model", ""))
    inventory = flatten_all_nested_fields(inventory)
    inventory = escape_multiline_strings(inventory)
    write_data_to_csv(inventory, csv_filename)

    table = PrettyTable()
    table.field_names = ["Index", "name", "mac", "model", "serial"]
    index_to_device = {}
    name_to_device = {}

    for idx, item in enumerate(inventory):
        table.add_row([idx, item.get("name", ""), item.get("mac", ""), item.get("model", ""), item.get("serial", "")])
        index_to_device[idx] = item
        name_to_device[item.get("name", "")] = item

    print(table)

    user_input = input("Enter the index or name of the device to view device: ").strip()

    # Try index
    if user_input.isdigit():
        idx = int(user_input)
        if idx in index_to_device:
            return index_to_device[idx].get("id")
        else:
            logging.warning("❌ Invalid index.")
            return None

    # Try name
    if user_input in name_to_device:
        return name_to_device[user_input].get("id")

    logging.warning("❌ Device not found by name or index.")
    return None

def show_site_device_inventory(site_id, device_type="all", csv_filename="SiteInventory.csv"):
    rawdata = mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id, type=device_type).data
    if not rawdata:
        print("No devices found for the selected site.")
        return

    inventory = sorted(rawdata, key=lambda x: x.get("model", ""))
    inventory = flatten_all_nested_fields(inventory)
    inventory = escape_multiline_strings(inventory)
    fields = get_all_unique_keys(inventory)
    write_data_to_csv(inventory, csv_filename)

    table = PrettyTable()
    table.field_names = fields

    if "model" in fields:
        try:
            table.sortby = "model"
        except Exception as e:
            logging.warning(f"⚠️ Could not sort table by 'model': {e}")

    for item in inventory:
        row = [item.get(field, "") for field in fields]
        table.add_row(row)

    logging.info("\n" + table.get_string())

def prompt_user_to_select_site_id_from_csv(csv_file="SiteList.csv"):
    """
    Prompts the user to select a site by index or name from SiteList.csv.
    Returns the corresponding site ID.
    """
    check_and_generate_csv(csv_file, export_org_site_list)

    with open(csv_file, mode='r', encoding='utf-8') as file:
        reader = list(csv.DictReader(file))
        index_to_site = {i: row for i, row in enumerate(reader)}
        name_to_site = {row["name"]: row for row in reader if "name" in row}

    print("\nAvailable Sites:")
    for idx, row in index_to_site.items():
        print(f"[{idx}] {row.get('name', 'Unnamed')}")

    user_input = input("\nEnter site index or name: ").strip()

    # Try index
    if user_input.isdigit():
        idx = int(user_input)
        if idx in index_to_site:
            site_id = index_to_site[idx].get("id")
            print(f"✅ Selected site: {index_to_site[idx].get('name')} (ID: {site_id})")
            return site_id
        else:
            print("❌ Invalid index.")
            return None

    # Try name
    if user_input in name_to_site:
        site_id = name_to_site[user_input].get("id")
        print(f"✅ Selected site: {user_input} (ID: {site_id})")
        return site_id

    print("❌ Site not found by name or index.")
    return None

def select_site():
    site_id = prompt_user_to_select_site_id_from_csv()
    if site_id:
        logging.info(f"✅ Selected site ID: {site_id}")
    else:
        logging.warning("❌ No site selected.")

def search_org_alarms():
    fetch_process_and_display_data(
        title="Search all Org Alarms:",
        api_call=mistapi.api.v1.orgs.alarms.searchOrgAlarms,
        filename="OrgAlarms.csv",
        limit=1000,
        duration="24h",
        status="open"
    )

def export_recent_device_events():
    logging.info("Search Org Device Events:")
    org_id = get_org_id()
    response = mistapi.api.v1.orgs.devices.searchOrgDeviceEvents(
        apisession, org_id, device_type="all", limit=1000, last_by="-24h"
    )
    rawdata = mistapi.get_all(response=response, mist_session=apisession)
    events = rawdata
    write_data_to_csv(events, "OrgDeviceEvents.csv")
    logging.info(json.dumps(events, indent=2))

def export_org_audit_logs():
    fetch_process_and_display_data(
        title="List Audit Logs:",
        api_call=mistapi.api.v1.orgs.logs.listOrgAuditLogs,
        filename="OrgAuditLogs.csv",
        limit=1000
    )

def export_org_site_list():
    fetch_process_and_display_data(
        title="Site List:",
        api_call=mistapi.api.v1.orgs.sites.searchOrgSites,
        filename="SiteList.csv",
        sort_key="name",  # or "site_id" if preferred
        limit=1000
    )

def export_org_device_inventory():
    fetch_process_and_display_data(
        title="Org Inventory:",
        api_call=mistapi.api.v1.orgs.inventory.getOrgInventory,
        filename="OrgInventory.csv",
        sort_key="model",
        limit=1000
    )

def export_org_device_statistics():
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
    print("Select a Site to View Device Inventory:")
    site_id = prompt_user_to_select_site_id_from_csv()
    if site_id:
        show_site_device_inventory(site_id)

def interactive_view_device_stats():
    """
    Prompts user to select a device and displays its detailed statistics.
    """
    interactive_device_action(
        fetch_function=mistapi.api.v1.sites.stats.getSiteDeviceStats,
        filename="DeviceStats.csv",
        description="Fetching detailed stats"
    )

def interactive_view_device_tests():
    """
    Prompts user to select a gateway device and displays its synthetic test stats.
    """
    interactive_device_action(
        fetch_function=mistapi.api.v1.sites.devices.getSiteDeviceSyntheticTest,
        filename="DeviceTestResults.csv",
        description="Fetching synthetic test stats",
        device_type="gateway"
    )

def interactive_view_device_config():
    """
    Prompts user to select a device and displays its configuration details.
    """
    interactive_device_action(
        fetch_function=mistapi.api.v1.sites.devices.getSiteDevice,
        filename="DeviceConfig.csv",
        description="Fetching device configuration"
    )

def export_all_org_devices():
    fetch_process_and_display_data(
        title="Org Devices:",
        api_call=mistapi.api.v1.orgs.devices.listOrgDevices,
        filename="OrgDevices.csv",
        sort_key="type",
        display_fields=["name", "mac"]
    )

def fetch_all_site_settings(apisession, org_id, limit=1000):
    print("Fetching all site settings...")

    # Use mistapi.get_all to ensure pagination is handled
    response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id)
    sites = mistapi.get_all(response=response, mist_session=apisession)

    all_configs = []
    for site in tqdm(sites, desc="Sites", unit="site"):
        site_id = site.get("id")
        site_name = site.get("name", "Unnamed Site")
        try:
            config = mistapi.api.v1.sites.setting.getSiteSetting(apisession, site_id).data
            config["site_id"] = site_id
            config["site_name"] = site_name
            all_configs.append(config)
        except Exception as e:
            print(f" ⚠️ Failed to fetch config for {site_name}: {e}")

    return all_configs

def export_all_site_settings():
    org_id = get_org_id()
    data = fetch_all_site_settings(apisession, org_id, limit=1000)
    if data:
        data = flatten_all_nested_fields(data)
        data = escape_multiline_strings(data)
        write_data_to_csv(data, "AllSiteConfigs.csv")
        logging.info("✅ Site configs saved to AllSiteConfigs.csv")
    else:
        logging.warning("⚠️ No site configs found.")

def export_all_gateway_device_configs():
    org_id = get_org_id()
    data = fetch_all_gateway_device_configs(apisession, org_id)
    if data:
        data = flatten_all_nested_fields(data)
        data = escape_multiline_strings(data)
        write_data_to_csv(data, "AllSiteGatewayConfigs.csv")
        logging.info("✅ Device configs saved to AllSiteGatewayConfigs.csv")
    else:
        logging.warning("⚠️ No device configs found.")

def fetch_all_gateway_device_configs(apisession, org_id):
    print("Fetching all sites in the org...")

    # Use mistapi.get_all to paginate through all sites
    response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id, limit=1000)
    sites = mistapi.get_all(response=response, mist_session=apisession)

    all_device_configs = []

    for site in tqdm(sites, desc="Sites", unit="site"):
        site_id = site.get("id")
        site_name = site.get("name", "Unnamed Site")

        try:
            # Use mistapi.get_all to paginate through all devices in the site
            response = mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id, type="gateway", limit=1000)
            devices = mistapi.get_all(response=response, mist_session=apisession)

            for device in tqdm(devices, desc=f"{site_name}", unit="device", leave=False):
                device_id = device.get("id")
                try:
                    config = mistapi.api.v1.sites.devices.getSiteDevice(apisession, site_id, device_id).data
                    config["site_id"] = site_id
                    config["site_name"] = site_name
                    all_device_configs.append(config)
                except Exception as e:
                    print(f"  ⚠️ Failed to fetch config for device {device_id} at {site_name}: {e}")
        except Exception as e:
            print(f"  ⚠️ Failed to list devices for site {site_name}: {e}")

    return all_device_configs

def export_nac_event_definitions():
    print("NAC Event Log Definitions:")
    rawdata = mistapi.api.v1.const.nac_events.listNacEventsDefinitions(apisession).data
    write_data_to_csv(rawdata, "NacEventDefinitions.csv")

def export_client_event_definitions():
    print("Client Event Log Definitions:")
    rawdata = mistapi.api.v1.const.client_events.listClientEventsDefinitions(apisession).data
    write_data_to_csv(rawdata, "ClientEventDefinitions.csv")

def export_device_event_definitions():
    print("Device Event Log Definitions:")
    rawdata = mistapi.api.v1.const.device_events.listDeviceEventsDefinitions(apisession).data
    write_data_to_csv(rawdata, "DeviceEventDefinitions.csv")

def export_mist_edge_event_definitions():
    print("Mist Edge Event Log Definitions:")
    rawdata = mistapi.api.v1.const.mxedge_events.listMxEdgeEventsDefinitions(apisession).data
    write_data_to_csv(rawdata, "MistEdgeEventDefinitions.csv")

def export_other_device_event_definitions():
    print("Other Event Log Definitions:")
    rawdata=(mistapi.api.v1.const.otherdevice_events.listOtherDeviceEventsDefinitions(apisession).data)
    write_data_to_csv(rawdata, "OtherEventDefinitions.csv")

def export_system_event_definitions():
    print("System Event Log Definitions:")
    rawdata=(mistapi.api.v1.const.system_events.listSystemEventsDefinitions(apisession).data)
    write_data_to_csv(rawdata, "SystemEventDefinitions.csv")

def export_alarm_definitions():
    print("Alarm Log Definitions:")
    rawdata=(mistapi.api.v1.const.alarm_defs.listAlarmDefinitions(apisession).data)
    alarm_defs = sorted(rawdata, key=lambda x: x.get("key", ""))
    write_data_to_csv(alarm_defs, "AlarmDefinitions.csv")
    table = PrettyTable()
    table.field_names = ["Key", "Display", "Group", "Severity", "Fields"]
    for alarm in alarm_defs:
        table.add_row([
            alarm.get("key"),
            alarm.get("display"),
            alarm.get("group"),
            alarm.get("severity"),
            ", ".join(alarm.get("fields", [])) if isinstance(alarm.get("fields"), list) else alarm.get("fields")
        ])

def export_all_gateway_synthetic_tests():
    logging.info("[INFO] Collecting synthetic test stats for all gateways in the org...")
    org_id = get_org_id()
    site_ids = get_sites_with_gateways(apisession, org_id)
    all_stats = []
    for site_id in tqdm(site_ids, desc="Sites", unit="site"):
        try:
            response = mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id, type="gateway")
            devices = mistapi.get_all(response=response, mist_session=apisession)
            for device in tqdm(devices, desc=f"Site {site_id}", unit="device", leave=False):
                device_id = device.get("id")
                try:
                    stats = mistapi.api.v1.sites.devices.getSiteDeviceSyntheticTest(apisession, site_id, device_id).data
                    stats["site_id"] = site_id
                    stats["site_name"] = device.get("site_name", "")
                    stats["device_id"] = device_id
                    stats["device_name"] = device.get("name", "")
                    all_stats.append(stats)
                except Exception as e:
                    logging.warning(f"⚠️ Failed to fetch test stats for device {device_id} at site {site_id}: {e}")
        except Exception as e:
            logging.warning(f"⚠️ Failed to list devices for site {site_id}: {e}")
    if all_stats:
        filename = "AllGatewaySyntheticTests.csv"
        flattened = flatten_all_nested_fields(all_stats)
        sanitized = escape_multiline_strings(flattened)
        write_data_to_csv(sanitized, filename)
        logging.info(f"✅ Synthetic test results saved to {filename}")
    else:
        logging.warning("⚠️ No synthetic test results found.")

def get_sites_with_gateways(apisession, org_id):
    print("[INFO] Fetching org inventory to find sites with gateways...")
    response = mistapi.api.v1.orgs.inventory.getOrgInventory(apisession, org_id, limit=1000)
    devices = mistapi.get_all(response=response, mist_session=apisession)
    gateway_sites = {device["site_id"] for device in devices if device.get("type") == "gateway" and "site_id" in device}
    print(f"[INFO] Found {len(gateway_sites)} sites with at least one gateway.")
    return list(gateway_sites)

def export_all_gateway_test_results_by_site():
    logging.info("[INFO] Searching all test results (including speed tests) for sites with gateways...")
    org_id = get_org_id()
    site_ids = get_sites_with_gateways(apisession, org_id)
    all_results = []

    if not site_ids:
        logging.warning("⚠️ No sites with gateways found.")
        return

    for site_id in tqdm(site_ids, desc="Sites", unit="site"):
        try:
            response = mistapi.api.v1.sites.synthetic_test.searchSiteSyntheticTest(
                apisession, site_id
            )
            if not hasattr(response, "data"):
                logging.warning(f"⚠️ No data attribute in response for site {site_id}")
                continue

            results = response.data.get("results", []) if isinstance(response.data, dict) else []
            logging.info(f"[{site_id}] Retrieved {len(results)} test results.")

            for result in results:
                result["site_id"] = site_id
                all_results.append(result)

        except Exception as e:
            logging.warning(f"⚠️ Failed to fetch test results for site {site_id}: {e}")

    if all_results:
        filename = "AllGatewayTestResults.csv"
        flattened = flatten_all_nested_fields(all_results)
        sanitized = escape_multiline_strings(flattened)
        write_data_to_csv(sanitized, filename)
        logging.info(f"✅ All test results saved to {filename}")
    else:
        logging.warning("⚠️ No test results found. CSV not created.")

def export_sites_with_location_info():
    logging.info("Listing Sites with Locations:")
    org_id = get_org_id()
    response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id)
    sites = mistapi.get_all(response=response, mist_session=apisession)
    site_data = []
    for site in sites:
        site_info = {
            "name": site.get("name", ""),
            "address": site.get("address", ""),
            "latitude": site.get("latlng", {}).get("lat", ""),
            "longitude": site.get("latlng", {}).get("lng", ""),
            "timezone": site.get("timezone", "")
        }
        site_data.append(site_info)
    site_data = escape_multiline_strings(site_data)
    write_data_to_csv(site_data, "SitesWithLocations.csv")
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
    logging.info("\n" + table.get_string())

def export_gateways_with_site_info():
    logging.info("Fetching Gateways with Site Info...")
    org_id = get_org_id()

    # Fetch site list
    site_response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id)
    sites = mistapi.get_all(response=site_response, mist_session=apisession)
    site_lookup = {
        site["id"]: {
            "name": site.get("name", ""),
            "address": site.get("address", "")
        } for site in sites
    }

    # Fetch org inventory
    inv_response = mistapi.api.v1.orgs.inventory.getOrgInventory(apisession, org_id)
    inventory = mistapi.get_all(response=inv_response, mist_session=apisession)

    def split_address(address):
        try:
            parts = address.split(", ")
            street = parts[0]
            city = parts[1]
            state_zip = parts[2].split()
            state = state_zip[0]
            zip_code = state_zip[1]
            country = parts[3]
            return street, city, state, zip_code, country
        except Exception:
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

    # Flatten and save
    gateways = flatten_all_nested_fields(gateways)
    gateways = escape_multiline_strings(gateways)
    gateways = sorted(gateways, key=lambda x: x.get("site_name", ""))
    write_data_to_csv(gateways, "GatewaysWithSiteInfo.csv")

    # Display
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
    logging.info("Fetching All Devices with Site Info...")
    org_id = get_org_id()

    site_response = mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id)
    sites = mistapi.get_all(response=site_response, mist_session=apisession)
    site_lookup = {
        site["id"]: {
            "name": site.get("name", ""),
            "address": site.get("address", "")
        } for site in sites
    }

    inv_response = mistapi.api.v1.orgs.inventory.getOrgInventory(apisession, org_id)
    inventory = mistapi.get_all(response=inv_response, mist_session=apisession)

    def split_address(address):
        try:
            parts = address.split(", ")
            street = parts[0]
            city = parts[1]
            state_zip = parts[2].split()
            state = state_zip[0]
            zip_code = state_zip[1]
            country = parts[3]
            return street, city, state, zip_code, country
        except Exception:
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

    enriched_devices = flatten_all_nested_fields(enriched_devices)
    enriched_devices = escape_multiline_strings(enriched_devices)
    enriched_devices = sorted(enriched_devices, key=lambda x: x.get("site_name", ""))
    write_data_to_csv(enriched_devices, "AllDevicesWithSiteInfo.csv")

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
    logging.info("\n" + table.get_string())

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
        check_and_generate_csv(filename, func, freshness_minutes=15)

    # Ensure SiteList.csv is generated before loading
    check_and_generate_csv('SiteList.csv', export_org_site_list, freshness_minutes=15)

    # Load the pulled data into dictionaries
    site_data = load_csv_into_dict('SiteList.csv', 'id')
    alarms_data = load_csv_into_dict('OrgAlarms.csv', 'site_id')
    events_data = load_csv_into_dict('OrgDeviceEvents.csv', 'site_id')
    devices_data = load_csv_into_dict('OrgDevices.csv', 'name')
    device_stats_data = load_csv_into_dict('OrgDeviceStats.csv', 'site_id')
    port_stats_data = load_csv_into_dict('OrgDevicePortStats.csv', 'site_id')

    if os.path.exists('AllGatewayTestResults.csv'):
        speedtest_data = load_csv_into_dict('AllGatewayTestResults.csv', 'site_id')
    else:
        logging.warning("⚠️ AllGatewayTestResults.csv not found. Skipping speedtest data.")
        speedtest_data = {}

    # Create a support package for each site with alarms or events
    for site_id, site_info in site_data.items():
        if not alarms_data.get(site_id) and not events_data.get(site_id):
            logging.info(f"Skipping site {site_id} — no alarms or events.")
            continue

        logging.info(f"Generating support package for site: {site_id}")
        support_data = {
            'alarms': alarms_data.get(site_id, []),
            'events': events_data.get(site_id, []),
            'devices': devices_data.get(site_id, []),
            'device_stats': device_stats_data.get(site_id, []),
            'port_stats': port_stats_data.get(site_id, []),
            'speedtests': speedtest_data.get(site_id, []),
        }

        support_package_filename = f"SupportPackage_{site_id}.csv"
        write_support_package_to_csv(support_data, support_package_filename)

    logging.info("✅ Support packages generated for applicable sites.")
    logging.info("✅ Support packages generated for all sites!")

def load_csv_into_dict(filename, key):
    # Load CSV data into a dictionary keyed by the specified column
    with open(filename, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)  # Create a CSV reader
        data_dict = {}  # Initialize an empty dictionary
        for row in reader:
            data_key = row[key]  # Get the value to use as the key
            if data_key not in data_dict:
                data_dict[data_key] = []  # Initialize a list for this key
            data_dict[data_key].append(row)  # Add the row to the dictionary
    return data_dict  # Return the dictionary

def write_support_package_to_csv(data, filename):
    # Write the support package data to a CSV file
    fieldnames = set()  # Initialize a set to collect all field names
    for section in data.values():
        for row in section:
            fieldnames.update(row.keys())  # Add all keys to the fieldnames set
    fieldnames = sorted(fieldnames)  # Sort the fieldnames

    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)  # Create a CSV writer
        writer.writeheader()  # Write the header row
        for section in data.values():
            for row in section:
                writer.writerow(row)  # Write each row to the CSV file
    logging.info(f"Support package written to {filename}")  # Log completion of the file write

def check_and_generate_csv(file_name, generate_function, freshness_minutes=15):
    """
    Checks if a CSV file exists and is fresh (modified within the last `freshness_minutes`).
    If not, it runs the `generate_function` to regenerate the file.
    """
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
    generate_function()

def poll_marvis_actions():
    print("🔍 Polling Marvis Actions...")
    org_id = get_org_id()

    # Call the Mist API to get Marvis actions
    response = mistapi.api.v1.orgs.troubleshoot.troubleshootOrg(apisession, org_id)
    rawdata = mistapi.get_all(response=response, mist_session=apisession)

    # Filter only open actions (state == "open")
    open_actions = [action for action in rawdata if action.get("state") == "open"]

    # Flatten and clean the data
    data = flatten_all_nested_fields(open_actions)
    data = escape_multiline_strings(data)

    # Write to CSV
    write_data_to_csv(data, "OpenMarvisActions.csv")
    print(f"✅ {len(open_actions)} open Marvis actions written to OpenMarvisActions.csv")

def export_current_guests():
    """
    Export all current guest users in the org to OrgCurrentGuests.csv
    """
    logging.info("Exporting all current guest users in the org...")
    org_id = get_org_id()
    response = mistapi.api.v1.orgs.guests.searchOrgGuestAuthorization(apisession, org_id, limit=1000)
    guests = mistapi.get_all(response=response, mist_session=apisession)
    guests = flatten_all_nested_fields(guests)
    guests = escape_multiline_strings(guests)
    write_data_to_csv(guests, "OrgCurrentGuests.csv")
    logging.info("✅ Current guests exported to OrgCurrentGuests.csv")

def export_historical_guests():
    """
    Export all guest users from the last 7 days to OrgHistoricalGuests.csv
    """
    logging.info("Exporting all guest users from the last 7 days...")
    org_id = get_org_id()
    # Calculate epoch for 7 days ago
    end_time = int(time.time())
    start_time = end_time - 7 * 24 * 3600
    response = mistapi.api.v1.orgs.guests.searchOrgGuestAuthorization(
        apisession, org_id, limit=1000, start=start_time, end=end_time
    )
    guests = mistapi.get_all(response=response, mist_session=apisession)
    guests = flatten_all_nested_fields(guests)
    guests = escape_multiline_strings(guests)
    write_data_to_csv(guests, "OrgHistoricalGuests.csv")
    logging.info("✅ Historical guests exported to OrgHistoricalGuests.csv")

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

        if not site_id or not device_id:
            continue

        try:
            # Get VC stats for this switch (returns a flat dict)
            vc_stats = mistapi.api.v1.sites.devices.getSiteDeviceVirtualChassis(apisession, site_id, device_id).data
            # Merge all switch info and VC info into a single dictionary
            entry = {**switch, **vc_stats}
            all_vc_stats.append(entry)
        except Exception as e:
            logging.warning(f"Failed to fetch VC stats for switch {name} ({device_id}): {e}")

    # Flatten and write to CSV
    all_vc_stats = flatten_all_nested_fields(all_vc_stats)
    all_vc_stats = escape_multiline_strings(all_vc_stats)
    write_data_to_csv(all_vc_stats, "OrgSwitchVCStats.csv")
    logging.info(PrettyTable(all_vc_stats))
    logging.info("✅ Switch VC stats exported to OrgSwitchVCStats.csv")

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
    "32": (export_all_switch_vc_stats, "Export all switch virtual chassis (VC/stacking) stats to CSV")
}

# --- CLI Argument Parsing ---
parser = argparse.ArgumentParser(description="MistHelper CLI Interface")
parser.add_argument("-O", "--org", help="Organization ID")
parser.add_argument("-M", "--menu", help="Menu option number to execute")
parser.add_argument("-S", "--site", help="Human-readable site name")
parser.add_argument("-D", "--device", help="Human-readable device name")
parser.add_argument("-P", "--port", help="Port ID")

args = parser.parse_args()

# If any CLI args are passed, override interactive mode
if len(sys.argv) > 1:
    if args.org:
        org_id = args.org  # Override global org_id

    # Resolve site name to site_id if needed
    if args.site:
        sites = mistapi.get_all(mistapi.api.v1.orgs.sites.listOrgSites(apisession, org_id), apisession)
        site_lookup = {site["name"]: site["id"] for site in sites}
        site_id = site_lookup.get(args.site)
        if not site_id:
            print(f"❌ Site name '{args.site}' not found.")
            sys.exit(1)
    else:
        site_id = None

    # Resolve device name to device_id if needed
    if args.device and site_id:
        devices = mistapi.get_all(mistapi.api.v1.sites.devices.listSiteDevices(apisession, site_id), apisession)
        device_lookup = {dev["name"]: dev["id"] for dev in devices}
        device_id = device_lookup.get(args.device)
        if not device_id:
            print(f"❌ Device name '{args.device}' not found at site '{args.site}'.")
            sys.exit(1)
    else:
        device_id = None

    # Execute the selected menu action
    if args.menu in menu_actions:
        func, _ = menu_actions[args.menu]
        func()
    else:
        print(f"❌ Invalid menu option: {args.menu}")
        sys.exit(1)

    sys.exit(0)  # Exit after CLI execution

# --- Interactive Menu Fallback ---
if len(sys.argv) == 1:
    print("\nAvailable Options:")
    for key, (func, description) in menu_actions.items():
        print(f"{key}: {description}")
    iwant = input("\nEnter your selection number now: ").strip()
    selected = menu_actions.get(iwant)
    if selected:
        func, _ = selected
        func()
    else:
        print("Invalid selection. Please try again.")
