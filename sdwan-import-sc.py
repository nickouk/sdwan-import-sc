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
keys = ['Device ID', 'System IP', 'Host Name', 'Site Id', 'Dual Stack IPv6 Default', 'Rollback Timer (sec)', 'basic_gpsl_longitude', 'basic_gpsl_latitude', 'provision_port_disable', 'vlan31_vrrp_pri', 'vlan31_vrrp_ipv4', 'vlan31_ipv4', 'vlan31_mask', 'vlan31_dhcp_net', 'vlan31_dhcp_mask', 'vlan31_dhcp_exclude', 'vlan31_dhcp_gateway', 'vlan120_vrrp_pri', 'vlan120_vrrp_ipv4', 'vlan120_ipv4', 'vlan120_mask', 'vlan120_dhcp_exclude', 'vlan100_vrrp_pri', 'vlan100_vrrp_ipv4', 'vlan100_ipv4', 'vlan100_mask', 'dhcp_6_basicConf_exclude', 'vlan40_vrrp_pri', 'vlan40_vrrp_ipv4', 'vlan40_ipv4', 'vlan40_mask', 'vlan40_dhcp_exclude', 'vlan30_vrrp_pri', 'vlan30_vrrp_ipv4', 'vlan30_ipv4', 'vlan30_mask', 'vlan30_dhcp_exclude', 'lan_vpn_100_nat_1_rangeStart', 'lan_vpn_100_nat_1_rangeEnd', 'lan_vpn_100_staticNat_1_translatedSourceIp', 'lan_vpn_100_staticNat_2_translatedSourceIp', 'loopback0_ipv4', 'loopback0_mask', 'vlan20_vrrp_pri', 'vlan20_vrrp_ipv4', 'vlan20_ipv4', 'vlan20_mask', 'vlan20_dhcp_net', 'vlan20_dhcp_mask', 'vlan20_dhcp_exclude', 'vlan20_dhcp_gateway', 'vlan10_vrrp_pri', 'vlan10_vrrp_ipv4', 'vlan10_ipv4', 'vlan10_mask', 'vlan10_dhcp_net', 'vlan10_dhcp_mask', 'vlan10_dhcp_exclude', 'vlan10_dhcp_gateway', 'vlan2_vrrp_pri', 'vlan2_vrrp_ipv4', 'vlan2_ipv4', 'vlan2_mask', 'vlan2_dhcp_net', 'vlan2_dhcp_mask', 'vlan2_dhcp_exclude', 'vlan2_dhcp_gateway', 'vlan80_vrrp_pri', 'vlan80_vrrp_ipv4', 'vlan80_ipv4', 'vlan80_mask', 'vlan70_vrrp_pri', 'vlan70_vrrp_ipv4', 'vlan70_ipv4', 'vlan70_mask', 'vlan60_vrrp_pri', 'vlan60_vrrp_ipv4', 'vlan60_ipv4', 'vlan60_mask', 'tloc_next_hop', 'tloc_bandwidth_up', 'tloc_bandwidth_down', 'wan_bandwidth_up', 'wan_bandwidth_down', 'wan_desc', 'ethpppoe_chapHost', 'ethpppoe_chapPwd', 'wan_color', 'ethpppoe_ipsecPrefer', 'wan_shapingRate', 'wan_track_addr']

vmanage_dict = {key: [] for key in keys}

# main loop - loop through the tracker sheet and build rows for the vmanage-import-sc.csv dictionary transforming some of the data

tracker_row = 3
postcode_list = []
print(f'{max_row} rows found ...\n')

while tracker_row <= max_row:

    # get the store number and pad to 4 digits
    store_num = str(tracker_sheet_obj.cell(row=tracker_row, column=1).value).zfill(4)

    # if store number is missing skip to next row
    if store_num == '0000' or store_num == 'None':
        tracker_row = tracker_row + 1
        continue

    # get the postcode
    postcode = str(tracker_sheet_obj.cell(row=tracker_row, column=3).value).upper().replace(' ', '')
    postcode_list.append(postcode)

    # get router 1 serial number
    router1_serial = str(tracker_sheet_obj.cell(row=tracker_row, column=4).value).upper()
    router1_serial = sanatise_serial(router1_serial)

    # get circuit 1 type and bandwidth
    circuit1_type = str(tracker_sheet_obj.cell(row=tracker_row, column=6).value).upper()
    circuit1_bw_down, circuit1_bw_up = circuit_bandwidth(circuit1_type)

    # get router 2 serial number
    router2_serial = str(tracker_sheet_obj.cell(row=tracker_row, column=11).value).upper()
    router2_serial = sanatise_serial(router2_serial)

    # get circuit 2 type and bandwidth
    circuit2_type = str(tracker_sheet_obj.cell(row=tracker_row, column=13).value).upper()
    circuit2_bw_down, circuit2_bw_up = circuit_bandwidth(circuit2_type)

    # generate store networks from store number
    print(store_num)
    store_net_oct2 = store_num[0:2]
    if store_net_oct2[0] == '0' or store_net_oct2[0] == '9':
        store_net_oct2 = store_net_oct2[1:]

    store_net_oct2 = int(store_net_oct2)
    store_net_oct3 = int(store_num[2:3])

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

    print("Store number: ", store_num)
    print("Postcode: ", postcode)
    print("Router 1 Serial: ", router1_serial)
    print("Circuit 1 Type: ", circuit1_type)
    print("Circuit 1 Bandwidth Up: ", circuit1_bw_up)
    print("Circuit 1 Bandwidth Down: ", circuit1_bw_down)
    print("Router 2 Serial: ", router2_serial)
    print("Circuit 2 Type: ", circuit2_type)
    print("Circuit 2 Bandwidth Up: ", circuit2_bw_up)
    print("Circuit 2 Bandwidth Down: ", circuit2_bw_down)
    print("VLAN60 IPv4 Network: ", vlan60_ipv4)
    print("VLAN70 IPv4 Network: ", vlan70_ipv4)
    print("VLAN80 IPv4 Network: ", vlan80_ipv4)
    print("VLAN10 IPv4 Network: ", vlan10_ipv4)
    print("VLAN20 IPv4 Network: ", vlan20_ipv4)
    print("VLAN30 IPv4 Network: ", vlan30_ipv4)
    print("VLAN31 IPv4 Network: ", vlan31_ipv4)
    print("VLAN40 IPv4 Network: ", vlan40_ipv4)
    print("VLAN100 IPv4 Network: ", vlan100_ipv4)
    print("VLAN120 IPv4 Network: ", vlan120_ipv4)
    print("")
    tracker_row = tracker_row + 1