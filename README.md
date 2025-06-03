This script pulls data and writes to a log file and also outputs data to multiple CSV depending on which option you select from the menu.

Use the sample .ENV file to setup your org specific ID and API token.

Flag Purpose Example -O Set org_id -O "abc123"

-M Menu option number -M 29

-S Human-readable site name (to be resolved to site_id) -S "Seattle HQ"

-D Human-readable device name (to be resolved to device_id) -D "GW-Edge-01"

-P Port ID -P "eth0"

Prerequisite imports:

import mistapi, csv, ast, json, time, logging, os, argparse, sys

from prettytable import PrettyTable from tqdm import tqdm from datetime import datetime, timedelta
