import re
from openpyxl import Workbook
import paramiko

# Match SSH command exactly
paramiko.Transport._preferred_kex = ('diffie-hellman-group1-sha1',)
paramiko.Transport._preferred_keys = ('ssh-rsa',)
paramiko.Transport._preferred_ciphers = ('aes128-cbc', '3des-cbc')
paramiko.Transport._preferred_macs = ('hmac-sha1', 'hmac-md5')

# Create SSH client
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

# Connect (same as ssh admin@172.20.10.5)
ssh.connect(
    hostname="172.20.10.5",
    username="admin",
    password="admin123",
)

# Run a test command
stdin, stdout, stderr = ssh.exec_command("show ip route")
def get_data(stdout):
    lines = stdout.readlines()

    # Patterns
    route_pattern = re.compile(r'^(S\*|C|L)\s+(\d+\.\d+\.\d+\.\d+)(/\d{1,2})')
    subnet_only_pattern = re.compile(r'^(\d+\.\d+\.\d+\.\d+)(/\d{1,2})')
    via_pattern = re.compile(r'via\s+(\d+\.\d+\.\d+\.\d+)')
    port_pattern = re.compile(r',\s*([A-Za-z0-9/]+)$')

    results = []

    last_type = None
    last_port = None

    for line in lines:
        line = line.strip()

        if not line or '-' in line:
            continue

        # Case 1: Normal route line (S*, C, L)
        match = route_pattern.search(line)

        if match:
            route_type = match.group(1)
            subnet = match.group(2)
            prefix = match.group(3)

            via_match = via_pattern.search(line)
            port_match = port_pattern.search(line)

            if via_match:
                port = via_match.group(1)
            elif port_match:
                port = port_match.group(1)
            else:
                port = "UNKNOWN"

            # Save last values
            last_type = route_type
            last_port = port

            results.append({
                "Type": route_type,
                "Subnet": subnet,
                "Prefix": prefix,
                "Port": port
            })

        # Case 2: Line starts with IP but no type → inherit previous
        else:
            subnet_match = subnet_only_pattern.search(line)

            if subnet_match and last_type and last_port:
                subnet = subnet_match.group(1)
                prefix = subnet_match.group(2)

                results.append({
                    "Type": last_type,
                    "Subnet": subnet,
                    "Prefix": prefix,
                    "Port": last_port
                })

    return results

def export_to_excel(results, output_file="Routers.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Task 2"

    ws.append(["Type", "Subnet", "Prefix", "Port"])

    for item in results:
        ws.append([item["Type"], item["Subnet"], item["Prefix"], item["Port"]])

    wb.save(output_file)

data = get_data(stdout)
export_to_excel(data)
ssh.close()