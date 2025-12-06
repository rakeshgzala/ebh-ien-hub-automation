# EBH iEN HUB Automation

## Authors

Anthony Sylvester [ansylves@cisco.com](mailto:ansylves@cisco.com)

## About

Verizon Wireless EBH iEN HUB Migration Automation script. This script is used to build the following files:
- Post-Migration Topology
- Pre-Checks
- Post-Checks
- Configuration Files:
  - EBH AL
  - HUB BL
  - HUB AL
  - DRAN SPOKE (Cisco)

- NOTE: The automation currently builds 'ssh server vrf CELL_MGMT'. __This will need to be removed after the primary issue gets resolved.__

## Requirements

- Python 3.10 or later (NOTE: Testing was done using Python 3.10)
  - Here is a link you can use: https://www.python.org/downloads/release/python-31011/
- A CIQ File that follows the format of the latest iEN Hub Migration CIQ Template
  - The file must have "CIQ" in the filename.
  - The file extension must be ".xlsx".
- Raw running configuration files from all hub devices. (Ciena configs not required.)
  - File names should be \<hostname>.cfg
  - Place the files inside the NP_DATA folder.
- PyCharm Community Edition or another IDE (NOTE: Use PyCharm if you want help when you run into issues.)
  - Here is the download link: https://www.jetbrains.com/pycharm/download/
  - NOTE: You may have to scroll down to select the Community Edition.

## How To Run

1. Download the code from the [GitHub Repo](https://wwwin-github.cisco.com/VzW-EBH/ebh-ien-hub-automation)
   1. Click on 'Code' and then copy the url. It should be: `https://wwwin-github.cisco.com/VzW-EBH/ebh-ien-hub-automation.git`
   2. Open PyCharm.
   3. Click on Git > Clone.
   4. Paste the URL and click Clone.
2. You will need to select an interpreter. Python 3.10 is recommended. (Python versions older than 3.10 are not supported.)
   1. Install any required Python Packages if creating a new interpreter in a virtual environment.
      1. These are typically:
         1. requests
         2. openpyxl
         3. ciscoconfparse
         4. networkx
         5. matplotlib
3. Place the CIQ file in the same directory as the script "ien_hub_config_builder.py".
   1. The CIQ file must have 'CIQ' in the file name, and the file extension must be '.xlsx'.
   2. It is assumed that CIQ Review has been completed and all data is correct and valid. NOTE: The script does not perform data validation!
4. Run the script:
   1. Open "ien_hub_config_builder.py"
   2. Run the file in debug mode.
      1. As the script works through each device in the CIQ, it will look in the NP_DATA directory to see if there is a legacy config. This would be saved as <Old Hostname>.cfg.
         1. If the legacy config was found (maybe you logged into the router & saved it manually) then the script will use that config file as needed. It might be used to carry over legacy interface config, l2vpn bridge domains, etc.
         2. If the legacy config was not found, it will attempt to pull the config for <Old Hostname> from NP:
            1. If the device was found in NP, the config will be saved to NP_DATA as <Old Hostname>.cfg and then the script will use that config file as needed.
            2. If you receive an error saying that the device's config wasn't found in NP, you will have to manually login to the router and save the running config as <Old Hostname>.cfg in the NP_DATA directory.
   3. If there are any WARNING or CRITICAL messages, review them and take an appropriate action. This may include updating CIQ sheet names, headers, or cells. If you update the CIQ, save it and re-run the script.
   4. After the script has completed, you should see 'Process finished with exit code 0'.
      1. If the code # is not 0, re-run the script. If you're getting the same exit code, then engage the script owner.
   5. The script will have created a directory named '<Site Name> - Files' with sub-folders, and files in those sub-folders. This is the output of the script.
   6. Review and validate that the output files are correct.

## Validation

Manual output validation is always required. Go through EVERY file.

__If the device used to be a hub router (h1/H2/etc.) then you MUST review the config output. It is highly likely that you will need to delete some carried over interfaces. More specifically, the interfaces that used to connect the legacy hub router to the L2 Aggregate device.__

## Caveats / Assumptions

1. It is assumed that all data in the CIQ is CORRECT. There is no data validation.
   1. The Cisco NCE should validate the data in the CIQ. See the CIQ Review section below for more guidance.
2. The automation supports a design where there is 1 connection for an EDN switch for the hub, and all hub devices will utilize this connection. If the customer deviates from this design, you will have to manually update the affected configs.
   1. If the customer has 2 EDN Switches, it's recommended that they trunk EDN2 to EDN1 and have EDN1 hang off a HUB AL. This way all devices in the HUB can be configured with a management IP from the same subnet.
3. If a DRAN Spoke is Ciena (has '-CN-' in the hostname) then the current config for that router is not required in NP_DATA, and of course we will not generate a config for that spoke.
4. The automation does not validate software version or installed SMUs (currently).
5. The automation does not account for GRE Spokes. That should be handled separately by the NCE.
6. The script does not carry over breakout config or breakout interfaces for HUB ALs.
7. The topology builder ONLY shows the connections to hub devices. If there are 2 EDN Switches and they are trunked together, the topology will only create the node for EDN SW 1 & the connection between EDN SW 1 & the HUB AL.
8. The NTP configuration of ALL devices (HUB BL, HUB AL, DRAN Spoke) is based off the ODD EBH AL. This is done to ensure standardization and correct configuration.
9. __The NCE must manually change the MTU for xNB and vDU interfaces to 1970.__

## CIQ Review - How to provide good data

Use the latest CIQ Template. It should be in ["Documents > CIQ Template" on our team SharePoint](https://cisco.sharepoint.com/sites/CX_VerizonEBH/Shared%20Documents/Forms/AllItems.aspx?newTargetListUrl=%2Fsites%2FCX%5FVerizonEBH%2FShared%20Documents&viewpath=%2Fsites%2FCX%5FVerizonEBH%2FShared%20Documents%2FForms%2FAllItems%2Easpx&id=%2Fsites%2FCX%5FVerizonEBH%2FShared%20Documents%2FCIQ%20Template&viewid=7f11d60e%2D360b%2D45de%2Daaa7%2D39d1b3c28df9).

1. All required sheet names should be present in the CIQ. See the latest CIQ template for reference. Capitalization & Punctuation matter.
2. All required headers should be present in the specific sheets in the CIQ. See the latest CIQ template for reference. Capitalization & Punctuation matter.
3. NOTE: The script stops processing data when it hits a blank row. It works this way for each required sheet.
4. Tab 'Hub Info':
   1. The data for 'Site Loopback0 Prefix' should come from prefix-set PRFX_EBH_LOOPBACKS on the EBH ALs. Put all network statements in CIDR notation, each on a new line. While you're in the cell, you can use alt+return to add a new line.
   2. The Prefix for the GRE VLANs can be found on the BE2x.1539 and BE2x.1540 interfaces. These are the /31 prefixes for GRE solutions and these connections exist between the MLS and the TAP/SAP BL. If there are 2 MLS pairs, list both /31 networks. Each pair will have 1 prefix for the ODD side and 1 for the EVEN side.
   3. All prefix entries should be in CIDR notation. Ex. 192.168.0.0/32
5. Tab 'Site Configuration Data':
   1. Make sure there are no extra spaces at the end of any cell.
   2. Verify that all IP addresses are unique.
   3. If the device is Greenfield (new device), then put 'N/A' in the 'Old Hostname' column.
   4. Verify connectivity to all devices via their Old MGMT IP before the migration. Make sure the Old Hostname matches exactly once you login to the Old MGMT IP.
   5. Approved eNSE SR device roles:
       1. EBH AL
       2. HUB BL
       3. HUB AL
       4. DRAN SPOKE
   6. Approved Models:
      1. EBH AL:
         1. NCS 55A1
      2. HUB BL:
         1. NCS 55A1
         2. NCS 5501
      3. HUB AL:
         1. NCS 540
         2. NCS 5501
      4. DRAN Spoke:
         1. NCS 540
         2. NCS 5501
         3. Other (If other vendor. Ex. Ciena may put cn5164)
   7. Verify that each device is running the latest approved SW Version (7.4.2) and all required SMUs are installed. Check NP if you need the list of SMUs.
   8. Verify that the New Hostname follows the naming standards (Applies to all except for EBH ALs. \<8-character CLLI\>\<3-character Role\>-\<2-character Vendor\>-\<9-digit Site ID\>-\<2-digit Device Number\>. Ex. GNBQNC15B4B-CI-062722489-01
   9. Verify that the management IPs for all hub devices are in the same /64 subnet.
      1. Check tab 'Non-RAN Services' and see which HUB-AL is hosting the EDN Switch, and then look at BVI40X for that HUB-AL.
         1. If the IPv6 address IS in the same subnet as the management IPs, you're good.
         2. If the IPv6 address IS NOT in the same subnet as the management IPs, then you need to ask Verizon to provide an IPv6 address from that subnet and assign it to BVI400 on that HUB-AL.
         3. Make sure that the management IP subnet is only owned by the EDN-hosting HUB-AL and not another HUB-AL.
      2. NOTE: Check to see if the current EDN-host device has BVI450 for UT equipment. If it does, that BVI450 is likely configured with an IPv4 address. Take the whole prefix and make sure you put that in 'BVI450 - IPv4 CIDR' for the new EDN-hosting HUB-AL.
   10. Make sure all Loopback IPs are unique.
   11. Verify that all Prefix SIDs are unique and not currently used. You can login to the EBH ALs and check the SR label table.
   12. Verify that the ISIS NET ID 'matches' the Loopback0 IPv4 address. See the LLD or MOP if you are unaware of how to check this.
   13. HUB-ALs must have BVI10X & BVI40X filled out. If the HUB-AL also has DSS connected, BVI15X will need to be filled out. See the MOP for more info.
   14. DRAN Spokes must have BVI100 & BVI400 filled out. If the DRAN Spoke also has DSS connected, BVI150 will need to be filled out. The existing DSS BVI is BVI310 - if configured.
   15. For HUB-ALs & DRAN Spokes, if there is UT equipment, you will carry over BVI450.
   16. HUB-ALs that are providing PTP-NB connectivity need to have BVI350 filled out.
   17. HUB-ALs & DRAN Spokes must have 'BVIs need suppress-ra?' set to 'Yes' if there is a Samsung VDU connected to the router. Enter 'No' if there isn't a Samsung VDU at the router.
   18. There should be 1 blank row at the end of the DRAN Spokes. This is how the automation knows to stop reading data.
       1. If the customer wants to add any info about other devices at the site, they should add this info below the blank row.
6. Tab 'EBH AL to HUB BL P2P':
   1. VLAN range is 501-900. If you receive a VLAN out of the range, ask Verizon why.
   2. The Interface columns should list the exact interface - not the sub-interface. Make sure the interface identifier is correct based on the platform. Ex. TEN0/0/0/1 if the connection is 10G and the platform is NCS 5501, or TEN0/0/0/1/0 if the connection is 10G and the platform is NCS 55A1.
   3. If the connection is dark fiber, the circuit rate should match the line rate.
   4. Verify that a Provider and Circuit ID are populated. These are used in the main interface description.
      1. NOTE: The sub-interface description will include the hostname and interface of the far end device.
      2. If Vz didn't populate these fields, just let them know it's used in the interface description. Get it in writing if they don't want to fill it out.
7. Tab 'HUB BL to HUB BL P2P':
   1. If this is filled out, the connection should use HUN0/0/0/34. If this is not HUN, ask Vz why. This is a connection within the hub and should be HUN to accommodate the traffic coming from the HUB-ALs.
8. Tab 'HUB BL to HUB AL P2P':
   1. Make sure each HUB-AL has a connection to each HUB-BL.
   2. HUB-AL should use HUN0/0/1/0 & HUN0/0/1/1.
9. Tab 'HUB AL to SPOKE P2P':
   1. VLAN range is 1001-4000. If you receive a VLAN out of the range, ask Verizon why.
   2. The Interface columns should list the exact interface - not the sub-interface. Make sure the interface identifier is correct based on the platform. Ex. TEN0/0/0/1 if the connection is 10G and the platform is NCS 5501, or TEN0/0/0/1/0 if the connection is 10G and the platform is NCS 55A1.
   3. If the connection is dark fiber, the circuit rate should match the line rate.
   4. Verify that a Provider and Circuit ID are populated. These are used in the main interface description.
      1. NOTE: The sub-interface description will include the hostname and interface of the far end device.
      2. If Vz didn't populate these fields, just let them know it's used in the interface description. Get it in writing if they don't want to fill it out.
10. Tab 'SPOKE to SPOKE P2P':
    1. 'Upstream SPOKE Hostname Connected to HUB AL' is the hostname of the SPOKE that is directly connected to the HUB-AL. This is how the script knows what BGP peers to configure for diasy-chain scenarios.
    2. VLAN range is 1001-4000. If you receive a VLAN out of the range, ask Verizon why.
    3. The Interface columns should list the exact interface - not the sub-interface. Make sure the interface identifier is correct based on the platform. Ex. TEN0/0/0/1 if the connection is 10G and the platform is NCS 5501, or TEN0/0/0/1/0 if the connection is 10G and the platform is NCS 55A1.
    4. If the connection is dark fiber, the circuit rate should match the line rate.
    5. Verify that a Provider and Circuit ID are populated. These are used in the main interface description.
       1. NOTE: The sub-interface description will include the hostname and interface of the far end device.
       2. If Vz didn't populate these fields, just let them know it's used in the interface description. Get it in writing if they don't want to fill it out.
11. Tab 'Non-RAN Services':
    1. The following should be provided:
       1. PTP Southbound (There might be 2 if there are 2 EGMs)
       2. PTP Northbound (There might be 2 if there are 2 EGMs)
       3. EDN Switch (There may be multiple)
       4. SiteBoss
    2. PTP Southbound should connect to the HUB-BLs
    3. PTP Northbound should connect to the HUB-ALs
    4. EDN Switch(es) should connect to the HUB-ALs
    5. There should be a blank row after the required data. The required data is the info outlined above. If the customer provides additional info, that needs to go below the blank row. Examples include info about which mgmt interface will connect to the different EDN ports, or even how 2 EDN switches will be trunked together.
12. All tabs:
    1. It's recommended to use either 2 characters for each interface name, or the full interface name. Ex. Hu or HundredGigE, Te or TenGigE.
