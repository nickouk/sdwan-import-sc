#!/usr/bin/env python3


# This code will be customer specific as each customer will have different template variables
#
# import tracker sheet and build template csv for import into vManage to cutdown on manual work required to deploy routers
# To adapt this code there are two main sections that require updating:
# Section 1 is the definition of vmanage_dict - each dictionary key maps to a column header which is a variable in a template
# Section 2 is the main loop that reads the tracker sheet, manipulates the data and then writes it into the dictionary
# Section 3 performs postcode lookups to obtain GPS coords and gathers a list of routes required for DNAC - the vmanage-import-[cust].csv file is written


# openpxyl is a library for handing MS Excel files

import openpyxl

# pandas is used for working with csv files

import pandas as pd

# requests allows API calls - used to correct the UK Postcodes which have no space

import requests

# allows whois lookup for public subnets ; static routes are required on DNAC

from ipwhois import IPWhois

# some standard libraries

import pprint
import ipaddress
import sys
import re
import math
import pickle
import json


def postcode_api(postcode_apilist):

    # Function for passing a list of postcodes to an external site for lookup
    # Correct the postcode format (missing space) and return long + lat values
    # Send API request, passing in postcodes as a list
    # NOTE Maximum 100 postcodes

    postcode_uri = 'https://api.postcodes.io/postcodes'

    # Raise an exception requests.HTTPException error is response is anything other than 200 (OK)

    json = {"postcodes": postcode_apilist}
    postcode_lookup = ''
    try:
        postcode_lookup = requests.post(
            postcode_uri,
            json={"postcodes": postcode_apilist}
        )
        postcode_lookup.raise_for_status()
    except requests.exceptions.ConnectionError:
        print(f'\nConnection error connecting to {postcode_uri}\nvManage import sheet has not been updated\n')
        sys.exit()
    except requests.HTTPError as error:
        print(f'\nHTTP Error:\n{error}')
        sys.exit()

    return (postcode_lookup)


def circuit_bandwidth(circuit_type):

    # Function to return circuit bandwidth based on circuit type

    if circuit_type == 'FTTP':
        return (80, 20)
    elif circuit_type == 'SOGEA':
        return (80, 20)
    elif circuit_type == 'FTTC':
        return (80, 20)
    elif circuit_type == 'ADSL':
        return (24, 1)
    elif circuit_type == 'OFNL Fibre':
        return (80, 20)
    else:
        # return a tuple so callers that unpack won't fail
        return (0, 0)


def sanatise_serial(serial):

    # Function to check the location code is 3 letters and remove leading 'S'
    if not serial:
        return ''
    # guard against short serials
    if len(serial) > 3:
        serial_check = serial[3]  # check 4th character is a number
    else:
        serial_check = ''
    if serial_check.isalpha() and serial[0].upper() == 'S':
        return (serial[1:].upper())
    else:
        return (serial.upper())

def wan_color(circuit_provider):

    # Function to return wan color based on circuit provider

    if circuit_provider == 'BT':
        return 'blue'
    elif circuit_provider == 'PXC':
        return 'green'
    elif circuit_provider == 'Other':
        return 'public-internet'
    else:
        return None

# -----------------------------
# --- Main code starts here ---
# -----------------------------
# Open the tracker sheet

tracker_wb_obj = openpyxl.load_workbook(
    '/mnt/c/Users/nick.oneill/Downloads/NOF2025 Rollout tracker.xlsx')
tracker_sheet_obj = tracker_wb_obj.active
# determine how many rows we have
max_row = tracker_sheet_obj.max_row

# initialise some variables
keys = ['Device ID', 'System IP', 'Host Name', 'Site Id', 'Dual Stack IPv6 Default', 'Rollback Timer (sec)', 'basic_gpsl_longitude', 'basic_gpsl_latitude', 'provision_port_disable', 'vlan31_vrrp_pri', 'vlan31_vrrp_ipv4', 'vlan31_ipv4', 'vlan31_mask', 'vlan31_dhcp_net', 'vlan31_dhcp_mask', 'vlan31_dhcp_exclude', 'vlan31_dhcp_gateway', 'vlan120_vrrp_pri', 'vlan120_vrrp_ipv4', 'vlan120_ipv4', 'vlan120_mask', 'vlan120_dhcp_exclude', 'vlan100_vrrp_pri', 'vlan100_vrrp_ipv4', 'vlan100_ipv4', 'vlan100_mask', 'vlan100_dhcp_exclude', 'vlan40_vrrp_pri', 'vlan40_vrrp_ipv4', 'vlan40_ipv4', 'vlan40_mask', 'vlan40_dhcp_exclude', 'vlan30_vrrp_pri', 'vlan30_vrrp_ipv4', 'vlan30_ipv4', 'vlan30_mask', 'vlan30_dhcp_exclude', 'lan_vpn_100_nat_1_rangeStart', 'lan_vpn_100_nat_1_rangeEnd', 'lan_vpn_100_staticNat_1_translatedSourceIp', 'lan_vpn_100_staticNat_2_translatedSourceIp', 'loopback0_ipv4', 'loopback0_mask', 'vlan20_vrrp_pri', 'vlan20_vrrp_ipv4', 'vlan20_ipv4', 'vlan20_mask', 'vlan20_dhcp_net', 'vlan20_dhcp_mask', 'vlan20_dhcp_exclude', 'vlan20_dhcp_gateway', 'vlan10_vrrp_pri', 'vlan10_vrrp_ipv4', 'vlan10_ipv4', 'vlan10_mask', 'vlan10_dhcp_net', 'vlan10_dhcp_mask', 'vlan10_dhcp_exclude', 'vlan10_dhcp_gateway', 'vlan2_vrrp_pri', 'vlan2_vrrp_ipv4', 'vlan2_ipv4', 'vlan2_mask', 'vlan2_dhcp_net', 'vlan2_dhcp_mask', 'vlan2_dhcp_exclude', 'vlan2_dhcp_gateway', 'vlan80_vrrp_pri', 'vlan80_vrrp_ipv4', 'vlan80_ipv4', 'vlan80_mask', 'vlan70_vrrp_pri', 'vlan70_vrrp_ipv4', 'vlan70_ipv4', 'vlan70_mask', 'vlan60_vrrp_pri', 'vlan60_vrrp_ipv4', 'vlan60_ipv4', 'vlan60_mask', 'tloc_next_hop', 'tloc_bandwidth_up', 'tloc_bandwidth_down', 'wan_bandwidth_up', 'wan_bandwidth_down', 'wan_desc', 'ethpppoe_chapHost', 'ethpppoe_chapPwd', 'wan_color', 'ethpppoe_ipsecPrefer', 'wan_shapingRate', 'wan_track_addr']

vmanage_dict = {key: [] for key in keys}

store_num_col = 1  # column A
store_type_col = 2  # column B
postcode_col = 4  # column D
router1_serial_col = 5  # column E
router1_mgmt_ip_col = 6  # column F
circuit1_provider_col =  7  # column G
circuit1_type_col = 8  # column H
circuit1_ref_col = 9  # column I
circuit1_ppp_name_col = 11  # column K
circuit1_ppp_pwd_col = 12  # column L
router2_serial_col = 13  # column M
router2_mgmt_ip_col = 14  # column N
circuit2_provider_col = 15  # column O
circuit2_type_col = 16  # column P
circuit2_ref_col = 17  # column Q
circuit2_ppp_name_col = 19  # column S
circuit2_ppp_pwd_col = 20  # column T
vlan2_col = 21  # column U

# main loop - loop through the tracker sheet and build rows for the vmanage-import-sc.csv dictionary transforming some of the data

tracker_row = 3
postcode_list = []
print(f'{max_row} rows found ...\n')

while tracker_row <= max_row:

    # get the store number and pad to 4 digits
    store_num = str(tracker_sheet_obj.cell(row=tracker_row, column=store_num_col).value).zfill(4)

    # if store number is missing skip to next row
    if store_num == '0000' or store_num == 'None':
        tracker_row = tracker_row + 1
        continue

    # get the store type
    store_type = str(tracker_sheet_obj.cell(row=tracker_row, column=store_type_col).value).upper()
    store_type = int(store_type[0])  # first character only
    site_id = f'{store_type}{store_num}'

    # get the postcode
    postcode = str(tracker_sheet_obj.cell(row=tracker_row, column=postcode_col).value).upper().replace(' ', '')
    postcode_list.append(postcode)

    # get router 1 serial number
    router1_serial = str(tracker_sheet_obj.cell(row=tracker_row, column=router1_serial_col).value).upper()
    router1_serial = sanatise_serial(router1_serial)

    # get circuit 1 type and bandwidth
    circuit1_type = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit1_type_col).value).upper()
    circuit1_bw_down, circuit1_bw_up = circuit_bandwidth(circuit1_type)
    circuit1_ref = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit1_ref_col).value).upper()

    # initialize router 2 variables in case of a single site router
    router2_serial = 'NONE'
    circuit2_type = 'NONE'
    circuit2_bw_down = 0
    circuit2_bw_up = 0
    circuit2_type = 'NONE'
    circuit2_ref = 'NONE'
    circuit2_ppp_name = 'NONE'
    circuit2_ppp_pwd = 'NONE'
    router2_mgmt_ip = 'NONE'
    router2_systemip = 'NONE'
    router2_hostname = 'NONE'
    router2_wan_color = 'NONE'

    # get router 2 serial number if present otherwsie assume a singe router site
    router2_serial = str(tracker_sheet_obj.cell(row=tracker_row, column=router2_serial_col).value).upper()
    if router2_serial != 'NONE':
        router2_serial = sanatise_serial(router2_serial)

        # get circuit 2 type and bandwidth
        circuit2_type = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit2_type_col).value).upper()
        circuit2_bw_down, circuit2_bw_up = circuit_bandwidth(circuit2_type)
        circuit2_ref = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit2_ref_col).value).upper()

        # get managment IP address for router 2
        router2_mgmt_ip = str(tracker_sheet_obj.cell(row=tracker_row, column=router2_mgmt_ip_col).value)
        if '/' not in router2_mgmt_ip: router2_mgmt_ip = router2_mgmt_ip + '/32'
        router2_mgmt_ip = ipaddress.ip_network(router2_mgmt_ip, strict=False)
        router2_systemip = router2_mgmt_ip.network_address

        # build router 2 hostname
        router2_hostname = f'SC-{store_type}-{store_num}-R2'

        # get provider for circuit 2
        circuit2_provider = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit2_provider_col).value).upper()
        router2_wan_color = wan_color(circuit2_provider)

        # get ppp name and password for circuit 2
        circuit2_ppp_name = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit2_ppp_name_col).value)
        circuit2_ppp_pwd = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit2_ppp_pwd_col).value)

        if circuit2_provider == "BT" and circuit2_ppp_name == "None":
            circuit2_ppp_name = 'dummy@bband1.com'
            circuit2_ppp_pwd = 'dummy'
        
        # we need to duplicate the postcode for router 2 which is not optimal but works for now
        postcode_list.append(postcode)

    # get managment IP address for router 1
    router1_mgmt_ip = str(tracker_sheet_obj.cell(row=tracker_row, column=router1_mgmt_ip_col).value)
    if '/' not in router1_mgmt_ip: router1_mgmt_ip = router1_mgmt_ip + '/32'
    router1_mgmt_ip = ipaddress.ip_network(router1_mgmt_ip, strict=False)
    router1_systemip = router1_mgmt_ip.network_address

    # build router 1 hostname
    router1_hostname = f'SC-{store_type}-{store_num}-R1'

    # get provider for circuit 1
    circuit1_provider = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit1_provider_col).value).upper()
    router1_wan_color = wan_color(circuit1_provider)

    # get ppp name and password for circuit 1
    circuit1_ppp_name = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit1_ppp_name_col).value)
    circuit1_ppp_pwd = str(tracker_sheet_obj.cell(row=tracker_row, column=circuit1_ppp_pwd_col).value)

    if circuit1_provider == "BT" and circuit1_ppp_name == "None":
        circuit1_ppp_name = 'dummy@bband1.com'
        circuit1_ppp_pwd = 'dummy'

    # get vlan 2 network
    vlan2_network = str(tracker_sheet_obj.cell(row=tracker_row, column=vlan2_col).value)
    if vlan2_network and '/' not in vlan2_network:
        vlan2_network = vlan2_network + '/29'
    vlan2_network = ipaddress.ip_network(vlan2_network, strict=False)

    # generate store networks from store number
    store_net_oct2 = store_num[0:2]
    if store_net_oct2[0] == '0' or store_net_oct2[0] == '9':
        store_net_oct2 = store_net_oct2[1:]

    store_net_oct2 = int(store_net_oct2)
    store_net_oct3 = int(store_num[2:4])

    vlan60_ipv4 = ipaddress.ip_network(f'151.{store_net_oct2}.{store_net_oct3}.0/24')
    vlan70_ipv4 = ipaddress.ip_network(f'10.{store_net_oct2}.{store_net_oct3}.0/24')
    vlan80_ipv4 = ipaddress.ip_network(f'192.168.100.0/24')
    #vlan2_ipv4 = ADD THIS LATER
    vlan10_ipv4 = ipaddress.ip_network(f'10.1{store_net_oct2}.{store_net_oct3}.0/25')
    vlan20_ipv4 = ipaddress.ip_network(f'10.1{store_net_oct2}.{store_net_oct3}.128/25')
    vlan30_ipv4 = ipaddress.ip_network(f'192.168.101.0/24')
    vlan31_ipv4 = ipaddress.ip_network(f'10.11{store_net_oct2}.{store_net_oct3}.0/28')
    vlan40_ipv4 = ipaddress.ip_network(f'192.168.102.0/24')
    vlan100_ipv4 = ipaddress.ip_network(f'192.168.103.0/24')
    vlan120_ipv4 = ipaddress.ip_network(f'192.168.104.0/24')

    # build the dictionary rows for router 1
    vmanage_dict['Device ID'].append("C1121X-8P-" + router1_serial)
    vmanage_dict['System IP'].append(str(router1_systemip))
    vmanage_dict['Host Name'].append(router1_hostname)
    vmanage_dict['Site Id'].append(site_id)
    vmanage_dict['Dual Stack IPv6 Default'].append('FALSE')
    vmanage_dict['Rollback Timer (sec)'].append('300')
    vmanage_dict['provision_port_disable'].append('FALSE')
    vmanage_dict['vlan31_vrrp_pri'].append('110')
    vmanage_dict['vlan31_vrrp_ipv4'].append(str(vlan31_ipv4.network_address + 14))
    vmanage_dict['vlan31_ipv4'].append(str(vlan31_ipv4.network_address + 12))
    vmanage_dict['vlan31_mask'].append(str(vlan31_ipv4.netmask))
    vmanage_dict['vlan31_dhcp_net'].append(str(vlan31_ipv4.network_address))
    vmanage_dict['vlan31_dhcp_mask'].append(str(vlan31_ipv4.netmask))
    vmanage_dict['vlan31_dhcp_exclude'].append(f'{str(vlan31_ipv4.network_address + 7)}-{str(vlan31_ipv4.network_address + 14)}')
    vmanage_dict['vlan31_dhcp_gateway'].append(str(vlan31_ipv4.network_address + 14))
    vmanage_dict['vlan120_vrrp_pri'].append('110')
    vmanage_dict['vlan120_vrrp_ipv4'].append(str(vlan120_ipv4.network_address + 254))
    vmanage_dict['vlan120_ipv4'].append(str(vlan120_ipv4.network_address + 252))
    vmanage_dict['vlan120_mask'].append(str(vlan120_ipv4.netmask))
    vmanage_dict['vlan120_dhcp_exclude'].append(f'{str(vlan120_ipv4.network_address + 128)}-{str(vlan120_ipv4.network_address + 254)}')
    vmanage_dict['vlan100_vrrp_pri'].append('110')
    vmanage_dict['vlan100_vrrp_ipv4'].append(str(vlan100_ipv4.network_address + 254))
    vmanage_dict['vlan100_ipv4'].append(str(vlan100_ipv4.network_address + 252))
    vmanage_dict['vlan100_mask'].append(str(vlan100_ipv4.netmask))
    vmanage_dict['vlan100_dhcp_exclude'].append(f'{str(vlan100_ipv4.network_address + 128)}-{str(vlan100_ipv4.network_address + 254)}')
    vmanage_dict['vlan40_vrrp_pri'].append('110')
    vmanage_dict['vlan40_vrrp_ipv4'].append(str(vlan40_ipv4.network_address + 254))
    vmanage_dict['vlan40_ipv4'].append(str(vlan40_ipv4.network_address + 252))
    vmanage_dict['vlan40_mask'].append(str(vlan40_ipv4.netmask))
    vmanage_dict['vlan40_dhcp_exclude'].append(f'{str(vlan40_ipv4.network_address + 128)}-{str(vlan40_ipv4.network_address + 254)}')
    vmanage_dict['vlan30_vrrp_pri'].append('110')
    vmanage_dict['vlan30_vrrp_ipv4'].append(str(vlan30_ipv4.network_address + 254))
    vmanage_dict['vlan30_ipv4'].append(str(vlan30_ipv4.network_address + 252))
    vmanage_dict['vlan30_mask'].append(str(vlan30_ipv4.netmask))
    vmanage_dict['vlan30_dhcp_exclude'].append(f'{str(vlan30_ipv4.network_address + 128)}-{str(vlan30_ipv4.network_address + 254)}')
    vmanage_dict['vlan20_vrrp_pri'].append('110')
    vmanage_dict['vlan20_vrrp_ipv4'].append(str(vlan20_ipv4.network_address + 126))
    vmanage_dict['vlan20_ipv4'].append(str(vlan20_ipv4.network_address + 124))
    vmanage_dict['vlan20_mask'].append(str(vlan20_ipv4.netmask))
    vmanage_dict['vlan20_dhcp_net'].append(str(vlan20_ipv4.network_address))
    vmanage_dict['vlan20_dhcp_mask'].append(str(vlan20_ipv4.netmask))
    vmanage_dict['vlan20_dhcp_exclude'].append(f'{str(vlan20_ipv4.network_address + 64)}-{str(vlan20_ipv4.network_address + 126)}')
    vmanage_dict['vlan20_dhcp_gateway'].append(str(vlan20_ipv4.network_address + 126))
    vmanage_dict['vlan10_vrrp_pri'].append('110')
    vmanage_dict['vlan10_vrrp_ipv4'].append(str(vlan10_ipv4.network_address + 126))
    vmanage_dict['vlan10_ipv4'].append(str(vlan10_ipv4.network_address + 124))
    vmanage_dict['vlan10_mask'].append(str(vlan10_ipv4.netmask))
    vmanage_dict['vlan10_dhcp_net'].append(str(vlan10_ipv4.network_address))
    vmanage_dict['vlan10_dhcp_mask'].append(str(vlan10_ipv4.netmask))
    vmanage_dict['vlan10_dhcp_exclude'].append(f'{str(vlan10_ipv4.network_address + 64)}-{str(vlan10_ipv4.network_address + 126)}')
    vmanage_dict['vlan10_dhcp_gateway'].append(str(vlan10_ipv4.network_address + 126))
    vmanage_dict['vlan60_vrrp_pri'].append('110')
    vmanage_dict['vlan60_vrrp_ipv4'].append(str(vlan60_ipv4.network_address + 254))
    vmanage_dict['vlan60_ipv4'].append(str(vlan60_ipv4.network_address + 252))
    vmanage_dict['vlan60_mask'].append(str(vlan60_ipv4.netmask))
    vmanage_dict['vlan70_vrrp_pri'].append('110')
    vmanage_dict['vlan70_vrrp_ipv4'].append(str(vlan70_ipv4.network_address + 254))
    vmanage_dict['vlan70_ipv4'].append(str(vlan70_ipv4.network_address + 252))
    vmanage_dict['vlan70_mask'].append(str(vlan70_ipv4.netmask))
    vmanage_dict['vlan80_vrrp_pri'].append('110')
    vmanage_dict['vlan80_vrrp_ipv4'].append(str(vlan80_ipv4.network_address + 254))
    vmanage_dict['vlan80_ipv4'].append(str(vlan80_ipv4.network_address + 252))
    vmanage_dict['vlan80_mask'].append(str(vlan80_ipv4.netmask))
    vmanage_dict['tloc_next_hop'].append(str('192.168.12.2'))
    vmanage_dict['tloc_bandwidth_up'].append(str(circuit1_bw_up * 100))
    vmanage_dict['tloc_bandwidth_down'].append(str(circuit1_bw_down * 100))
    vmanage_dict['wan_bandwidth_up'].append(str(circuit1_bw_up * 100))
    vmanage_dict['wan_bandwidth_down'].append(str(circuit1_bw_down * 100))
    vmanage_dict['wan_desc'].append(f'{circuit1_ref} - {circuit1_type} via {circuit1_provider}')
    vmanage_dict['ethpppoe_chapHost'].append(circuit1_ppp_name)
    vmanage_dict['ethpppoe_chapPwd'].append(circuit1_ppp_pwd)
    vmanage_dict['wan_color'].append(router1_wan_color)
    vmanage_dict['ethpppoe_ipsecPrefer'].append('111')
    vmanage_dict['wan_shapingRate'].append(circuit1_bw_up * 100)
    vmanage_dict['wan_track_addr'].append('1.1.1.1')

    # if we have a router 2 build the dictionary rows for router 2
    if router2_serial != 'NONE':
        vmanage_dict['Device ID'].append("C1121X-8P-" + router2_serial)
        vmanage_dict['System IP'].append(str(router2_systemip))
        vmanage_dict['Host Name'].append(router2_hostname)
        vmanage_dict['Site Id'].append(site_id)
        vmanage_dict['Dual Stack IPv6 Default'].append('FALSE')
        vmanage_dict['Rollback Timer (sec)'].append('300')
        vmanage_dict['provision_port_disable'].append('FALSE')
        vmanage_dict['vlan31_vrrp_pri'].append('100')
        vmanage_dict['vlan31_vrrp_ipv4'].append(str(vlan31_ipv4.network_address + 14))
        vmanage_dict['vlan31_ipv4'].append(str(vlan31_ipv4.network_address + 13))
        vmanage_dict['vlan31_mask'].append(str(vlan31_ipv4.netmask))
        vmanage_dict['vlan31_dhcp_net'].append(str(vlan31_ipv4.network_address))
        vmanage_dict['vlan31_dhcp_mask'].append(str(vlan31_ipv4.netmask))
        vmanage_dict['vlan31_dhcp_exclude'].append(f'{str(vlan31_ipv4.network_address + 1)}-{str(vlan31_ipv4.network_address + 6)}";"{str(vlan31_ipv4.network_address + 12)}-{str(vlan31_ipv4.network_address + 14)}')
        vmanage_dict['vlan31_dhcp_gateway'].append(str(vlan31_ipv4.network_address + 14))
        vmanage_dict['vlan120_vrrp_pri'].append('100')
        vmanage_dict['vlan120_vrrp_ipv4'].append(str(vlan120_ipv4.network_address + 254))
        vmanage_dict['vlan120_ipv4'].append(str(vlan120_ipv4.network_address + 253))
        vmanage_dict['vlan120_mask'].append(str(vlan120_ipv4.netmask))
        vmanage_dict['vlan120_dhcp_exclude'].append(f'{str(vlan120_ipv4.network_address + 1)}-{str(vlan120_ipv4.network_address + 127)}";"{str(vlan120_ipv4.network_address + 252)}-{str(vlan120_ipv4.network_address + 254)}')
        vmanage_dict['vlan100_vrrp_pri'].append('100')
        vmanage_dict['vlan100_vrrp_ipv4'].append(str(vlan100_ipv4.network_address + 254))
        vmanage_dict['vlan100_ipv4'].append(str(vlan100_ipv4.network_address + 253))
        vmanage_dict['vlan100_mask'].append(str(vlan100_ipv4.netmask))
        vmanage_dict['vlan100_dhcp_exclude'].append(f'{str(vlan100_ipv4.network_address + 1)}-{str(vlan100_ipv4.network_address + 127)}";"{str(vlan100_ipv4.network_address + 252)}-{str(vlan100_ipv4.network_address + 254)}')
        vmanage_dict['vlan40_vrrp_pri'].append('100')
        vmanage_dict['vlan40_vrrp_ipv4'].append(str(vlan40_ipv4.network_address + 254))
        vmanage_dict['vlan40_ipv4'].append(str(vlan40_ipv4.network_address + 253))
        vmanage_dict['vlan40_mask'].append(str(vlan40_ipv4.netmask))
        vmanage_dict['vlan40_dhcp_exclude'].append(f'{str(vlan40_ipv4.network_address + 1)}-{str(vlan40_ipv4.network_address + 127)}";"{str(vlan40_ipv4.network_address + 252)}-{str(vlan40_ipv4.network_address + 254)}')
        vmanage_dict['vlan30_vrrp_pri'].append('100')
        vmanage_dict['vlan30_vrrp_ipv4'].append(str(vlan30_ipv4.network_address + 254))
        vmanage_dict['vlan30_ipv4'].append(str(vlan30_ipv4.network_address + 253))
        vmanage_dict['vlan30_mask'].append(str(vlan30_ipv4.netmask))
        vmanage_dict['vlan30_dhcp_exclude'].append(f'{str(vlan30_ipv4.network_address + 1)}-{str(vlan30_ipv4.network_address + 127)}";"{str(vlan30_ipv4.network_address + 252)}-{str(vlan30_ipv4.network_address + 254)}')
        vmanage_dict['vlan20_vrrp_pri'].append('100')
        vmanage_dict['vlan20_vrrp_ipv4'].append(str(vlan20_ipv4.network_address + 126))
        vmanage_dict['vlan20_ipv4'].append(str(vlan20_ipv4.network_address + 125))
        vmanage_dict['vlan20_mask'].append(str(vlan20_ipv4.netmask))
        vmanage_dict['vlan20_dhcp_net'].append(str(vlan20_ipv4.network_address))
        vmanage_dict['vlan20_dhcp_mask'].append(str(vlan20_ipv4.netmask))
        vmanage_dict['vlan20_dhcp_exclude'].append(f'{str(vlan20_ipv4.network_address + 1)}-{str(vlan20_ipv4.network_address + 63)}";"{str(vlan20_ipv4.network_address + 124)}-{str(vlan20_ipv4.network_address + 126)}')
        vmanage_dict['vlan20_dhcp_gateway'].append(str(vlan20_ipv4.network_address + 126))
        vmanage_dict['vlan10_vrrp_pri'].append('100')
        vmanage_dict['vlan10_vrrp_ipv4'].append(str(vlan10_ipv4.network_address + 126))
        vmanage_dict['vlan10_ipv4'].append(str(vlan10_ipv4.network_address + 125))
        vmanage_dict['vlan10_mask'].append(str(vlan10_ipv4.netmask))
        vmanage_dict['vlan10_dhcp_net'].append(str(vlan10_ipv4.network_address))
        vmanage_dict['vlan10_dhcp_mask'].append(str(vlan10_ipv4.netmask))
        vmanage_dict['vlan10_dhcp_exclude'].append(f'{str(vlan10_ipv4.network_address + 1)}-{str(vlan10_ipv4.network_address + 63)}";"{str(vlan10_ipv4.network_address + 124)}-{str(vlan10_ipv4.network_address + 126)}')
        vmanage_dict['vlan10_dhcp_gateway'].append(str(vlan10_ipv4.network_address + 126))
        vmanage_dict['vlan60_vrrp_pri'].append('100')
        vmanage_dict['vlan60_vrrp_ipv4'].append(str(vlan60_ipv4.network_address + 254))
        vmanage_dict['vlan60_ipv4'].append(str(vlan60_ipv4.network_address + 253))
        vmanage_dict['vlan60_mask'].append(str(vlan60_ipv4.netmask))
        vmanage_dict['vlan70_vrrp_pri'].append('100')
        vmanage_dict['vlan70_vrrp_ipv4'].append(str(vlan70_ipv4.network_address + 254))
        vmanage_dict['vlan70_ipv4'].append(str(vlan70_ipv4.network_address + 253))
        vmanage_dict['vlan70_mask'].append(str(vlan70_ipv4.netmask))
        vmanage_dict['vlan80_vrrp_pri'].append('100')
        vmanage_dict['vlan80_vrrp_ipv4'].append(str(vlan80_ipv4.network_address + 254))
        vmanage_dict['vlan80_ipv4'].append(str(vlan80_ipv4.network_address + 253))
        vmanage_dict['vlan80_mask'].append(str(vlan80_ipv4.netmask))
        vmanage_dict['tloc_next_hop'].append(str('192.168.21.1'))
        vmanage_dict['tloc_bandwidth_up'].append(str(circuit2_bw_up * 100))
        vmanage_dict['tloc_bandwidth_down'].append(str(circuit2_bw_down * 100))
        vmanage_dict['wan_bandwidth_up'].append(str(circuit2_bw_up * 100))
        vmanage_dict['wan_bandwidth_down'].append(str(circuit2_bw_down * 100))
        vmanage_dict['wan_desc'].append(f'{circuit2_ref} - {circuit2_type} via {circuit2_provider}')
        vmanage_dict['ethpppoe_chapHost'].append(circuit2_ppp_name)
        vmanage_dict['ethpppoe_chapPwd'].append(circuit2_ppp_pwd)
        vmanage_dict['wan_color'].append(router2_wan_color)
        vmanage_dict['ethpppoe_ipsecPrefer'].append('0')
        vmanage_dict['wan_shapingRate'].append(circuit2_bw_up * 100)
        vmanage_dict['wan_track_addr'].append('8.8.8.8')

    tracker_row = tracker_row + 1

# end of main loop
# perform postcode lookups to obtain GPS coords
print('Performing postcode lookups ...\n')

# Break the postcode list into chunks of 100 as the API has a max 100 limit

latlist = []
longlist = []

while len(postcode_list) > 100:
    postcode_max100 = postcode_list[0:100]
    postcode_result = postcode_api(postcode_max100)
    postcode_df = pd.json_normalize(postcode_result.json()['result'],sep='_')
    # update the csv dictionary with the lat and long values returned by the API
    latlist = latlist + (postcode_df['result_latitude'].to_list())
    longlist = longlist+ (postcode_df['result_longitude'].to_list())
    # remove the postcodes we have lookedup from the list
    postcode_list = postcode_list[100:]

if len(postcode_list) > 0:
    postcode_result = postcode_api(postcode_list)
    postcode_df = pd.json_normalize(postcode_result.json()['result'],sep='_')
    latlist = latlist + (postcode_df['result_latitude'].to_list())
    longlist = longlist+ (postcode_df['result_longitude'].to_list())
    
# update the csv dictionary with the lat and long values returned by the API
vmanage_dict['basic_gpsl_latitude'] = latlist
vmanage_dict['basic_gpsl_longitude'] = longlist

print(json.dumps(vmanage_dict, indent=1))

# create the dataframe from the dictionary we built
df = pd.DataFrame(vmanage_dict)

# write the dataframe to a csv ready for import into vManage
df.to_csv('/mnt/c/Users/nick.oneill/Downloads/vmanage-import-sc.csv', index=False)

# all done
print('\nvmanage-import-sc.csv has been created :)\n')