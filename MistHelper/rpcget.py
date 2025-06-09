import os
import json
from lxml import etree
from jnpr.junos import Device
from jnpr.junos.exception import ConnectError
from dotenv import load_dotenv

# Load credentials from .env file
load_dotenv()
rpcuser = os.getenv('rpcuser')
rpcpassword = os.getenv('rpcpassword')
rpchost = os.getenv('rpchost')

def extract_show_commands():
    try:
        dev = Device(host=rpchost, user=rpcuser, passwd=rpcpassword)
        dev.open()

        # Run the CLI command and get XML output
        response = dev.rpc.cli("""help apropos " "| display xml | no-more""", format="xml")
        root = response

        # Extract <output> tags
        outputs = root.xpath('//output')
        commands = {}
        i = 0
        while i < len(outputs) - 1:
            cmd = outputs[i].text.strip()
            desc = outputs[i + 1].text.strip()
            commands[cmd] = desc
            i += 2

        # Save to JSON
        with open("show_command_help.json", "w", encoding="utf-8") as f:
            json.dump(commands, f, indent=4)

        print("✅ show_command_help.json created successfully.")
        dev.close()

    except ConnectError as e:
        print(f"❌ Failed to connect to device: {e}")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    extract_show_commands()
