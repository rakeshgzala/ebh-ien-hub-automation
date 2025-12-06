'''
#############################################################################
### This script was created for use with the Verizon Wireless EBH team.   ###
### Authors:                                                              ###
### Rakesh Zala                                                           ###
###                                                                       ###
#############################################################################
'''

import re, sys, os, json, logging, io, copy, math, pip
from datetime import datetime, timedelta
from openpyxl import load_workbook                  # Used to read the Excel CIQ file
from mimir import Mimir, MimirAuthenticationError   # Mimir is used to pull running configs from Network Profiler.
from ciscoconfparse import CiscoConfParse           # CiscoConfParse is used to parse the running configs.
import networkx as nx                   # Used for building the topology connections
import matplotlib                       # Used for plotting the topology
import matplotlib.pyplot as plt         # Used for displaying the topology
# Import the telemetry module for CX Catalog. If the required module is not installed, install it.
try:
    import aide
except:
    import sys
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'git+https://wwwin-github.cisco.com/AIDE/aide-python-agent.git'])
    import aide


# Create a custom logger at the global level
logger = logging.getLogger(__name__)
# Create handlers
c_handler = logging.StreamHandler()
f_handler_debug = logging.FileHandler('debug-all.log', mode='w+')
f_handler_console = logging.FileHandler('debug-console.log', mode='w+')
c_handler.setLevel(logging.WARNING)
f_handler_debug.setLevel(logging.DEBUG)
f_handler_console.setLevel(logging.WARNING)
# Create formatters and add it to handlers
c_format = logging.Formatter('[%(filename)s:%(lineno)5s - %(funcName)20s ] - %(levelname)s - %(message)s')
f_format = logging.Formatter('%(asctime)s - [%(filename)s:%(lineno)5s - %(funcName)20s ] - %(levelname)s - %(message)s')
c_handler.setFormatter(c_format)
f_handler_debug.setFormatter(f_format)
f_handler_console.setFormatter(f_format)
# Add handlers to the logger
logger.addHandler(c_handler)
logger.addHandler(f_handler_debug)
logger.addHandler(f_handler_console)
# Set logger to DEBUG level so that all logs get passed to the handlers.
logger.setLevel(logging.DEBUG)


class build_ebh_al:

    def check(self, now, host_data):

        def pre_check():
            conf = f'''### Pre-Checks ###
end
copy running-config harddisk:PreRunCfgBkp_{now.strftime("%m%d%Y")}.cfg

terminal length 0
show bfd session
show isis adjacency
show isis segment-routing label table
show isis topology
show route vrf all
show route vrf RAN ipv6
show route vrf CELL_MGMT ipv6
show route vrf all ipv6
show arp vrf all
show ipv6 neighbors
show bgp l2vpn evpn summary
show bgp l2vpn evpn
show bgp ipv4 label summary
show bgp ipv4 label
show bgp labels'''
            for intf in host_data['Interfaces']:
                conf += f'''
show interface {intf}
show controller {intf.split(".")[0]}'''
            conf += f'''
show calendar
show ntp status
show ntp associations detail
show policy-map interface all | include "input|output"
show ptp interfaces brief
show ptp advertised-clock
show frequency synchronization summary
'''
            return conf

        def post_check():
            conf = f'''### Post-Checks ###
end
terminal length 0
show bfd session
show isis adjacency
show isis segment-routing label table
show isis topology
show route vrf all
show route vrf RAN ipv6
show route vrf CELL_MGMT ipv6
show route vrf all ipv6
show arp vrf all
show ipv6 neighbors
show bgp l2vpn evpn summary
show bgp l2vpn evpn
show bgp ipv4 label summary
show bgp ipv4 label
show bgp labels'''
            for intf in host_data['Interfaces']:
                conf += f'''
show interface {intf}
show controller {intf.split(".")[0]}'''
            conf += f'''
show calendar
show ntp status
show ntp associations detail
show policy-map interface all | include "input|output"
show ptp interfaces brief
show ptp advertised-clock
show frequency synchronization summary
'''
            return conf

        logger.debug(f"Class: {self.__class__.__name__}")
        conf = pre_check() + "\n" + post_check()
        return conf

    def config(self, host_data, ciq_db, runfile):

        def ports():
            conf = f'''### Port Config to HUB BL ###
configure terminal'''
            for intf in host_data["Interfaces"]:
                if "HUB BL" in ciq_db[host_data["Interfaces"][intf]["description"].split('_')[0]]["eNSESR Role"]:
                    # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                    intf_sh = intf[:2].lower()
                    if "Breakout" in host_data["Interfaces"][intf].keys():
                        if f'''controller optics {host_data["Interfaces"][intf]["Breakout"]}''' not in conf:
                            conf += f'''
root
controller optics {host_data["Interfaces"][intf]["Breakout"]}
 breakout 4x10
commit
'''
                    circuit_rate = int(host_data["Interfaces"][intf]["circuit_rate"])
                    if intf_sh == "gi":
                        if circuit_rate <= 0 or circuit_rate > 1000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 1000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 1000:
                            shaper_config = runfile.find_objects(f'''policy-map {circuit_rate}MB-NE''')
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in shaper_config:
                                conf += f'''
root
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 end-policy-map
'''
                        elif circuit_rate == 1000:
                            intf_bw = 1000000
                            egress_policy = f'''QUEUES-OUT'''
                    elif intf_sh == "te":
                        if circuit_rate <= 0 or circuit_rate > 10000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 10000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 10000:
                            shaper_config = runfile.find_objects(f'''policy-map {circuit_rate}MB-NE''')
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in shaper_config:
                                conf += f'''
root
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 end-policy-map
'''
                        elif circuit_rate == 10000:
                            intf_bw = 10000000
                            egress_policy = f'''QUEUES-OUT'''
                    elif intf_sh == "hu":
                        if circuit_rate <= 0 or circuit_rate > 100000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 100000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 100000:
                            shaper_config = runfile.find_objects(f'''policy-map {circuit_rate}MB-NE''')
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in shaper_config:
                                conf += f'''
root
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 end-policy-map
'''
                        elif circuit_rate == 100000:
                            intf_bw = 100000000
                            egress_policy = f'''QUEUES-OUT'''
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                        exit()
                    conf += f'''
root
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["cid"]}
 mtu 9100
 load-interval 30
 carrier-delay up 50 down 25
interface {intf}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 bandwidth {intf_bw}
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
 load-interval 30
 encapsulation dot1q {host_data["Interfaces"][intf]["VLAN"]}
'''
            conf += f'''
commit show-error
'''
            return conf

        def isis():
            conf = f'''
### ISIS Process 5 ###
root
router isis 5'''
            for intf in host_data["Interfaces"]:
                conf += f'''
 interface {intf}
  circuit-type level-1
  bfd minimum-interval 50
  bfd multiplier 5
  bfd fast-detect ipv4
  point-to-point
  hello-padding disable
  hello-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
  address-family ipv4 unicast
   fast-reroute per-prefix
   fast-reroute per-prefix ti-lfa
   metric 1000000
'''
            conf += f'''
commit show-error
'''
            return conf

        def bgp():
            conf = f'''
### BGP ###
root
router bgp {host_data["COIN ASN"]}'''
            for peer in host_data["BGP_Peers"]:
                conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_BL
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
'''
            conf += f'''
commit show-error
'''
            return conf

        logger.debug(f"Class: {self.__class__.__name__}")
        conf = ""
        conf += ports()
        conf += isis()
        conf += bgp()
        return conf


class build_hub_bl:

    def check(self, now, host_data):

        def pre_check():
            conf = f'''### Pre-Checks ###
end
copy running-config harddisk:PreRunCfgBkp_{now.strftime("%m%d%Y")}.cfg

terminal length 0
show bfd session
show isis adjacency
show isis segment-routing label table
show isis topology
show route vrf all
show route vrf RAN ipv6
show route vrf CELL_MGMT ipv6
show route vrf all ipv6
show arp vrf all
show ipv6 neighbors
show bgp l2vpn evpn summary
show bgp l2vpn evpn
show bgp ipv4 label summary
show bgp ipv4 label
show bgp labels'''
            for intf in host_data['Interfaces']:
                conf += f'''
show interface {intf}
show controller {intf.split(".")[0]}'''
            conf += f'''
show calendar
show ntp status
show ntp associations detail
show policy-map interface all | include "input|output"
show ptp interfaces brief
show ptp advertised-clock
show frequency synchronization summary
'''
            return conf

        def post_check():
            conf = f'''### Post-Checks ###
end
terminal length 0
show bfd session
show isis adjacency
show isis segment-routing label table
show isis topology
show route vrf all
show route vrf RAN ipv6
show route vrf CELL_MGMT ipv6
show route vrf all ipv6
show arp vrf all
show ipv6 neighbors
show bgp l2vpn evpn summary
show bgp l2vpn evpn
show bgp ipv4 label summary
show bgp ipv4 label
show bgp labels'''
            for intf in host_data['Interfaces']:
                conf += f'''
show interface {intf}
show controller {intf.split(".")[0]}'''
            conf += f'''
show calendar
show ntp status
show ntp associations detail
show policy-map interface all | include "input|output"
show ptp interfaces brief
show ptp advertised-clock
show frequency synchronization summary
'''
            return conf

        logger.debug(f"Class: {self.__class__.__name__}")
        conf = pre_check() + "\n" + post_check()
        return conf

    def config(self, host_data, ciq_db, runfile, mtso_config):

        def admin_part_1():
            conf = ""
            # MOP Section: Admin Part 1
            conf += f'''
hostname {host_data["New Hostname"]}
taskgroup READ-ONLY-TGRP
 task read fr
 task read li
 task read aaa
 task read acl
 task read atm
 task read bfd
 task read bgp
 task read cdp
 task read cef
 task read cgn
 task read eem
 task read ppp
 task read qos
 task read rib
 task read rip
 task read sbc
 task read ancp
 task read bcdl
 task read boot
 task read diag
 task read dwdm
 task read hdlc
 task read hsrp
 task read ipv4
 task read ipv6
 task read isis
 task read lpts
 task read ospf
 task read ouni
 task read snmp
 task read vlan
 task read vrrp
 task read admin
 task read eigrp
 task read l2vpn
 task read bundle
 task read crypto
 task read fabric
 task read static
 task read sysmgr
 task read system
 task read tunnel
 task read drivers
 task read logging
 task read monitor
 task read mpls-te
 task read netflow
 task read network
 task read pos-dpt
 task read firewall
 task read mpls-ldp
 task read pkg-mgmt
 task read fault-mgr
 task read interface
 task read inventory
 task read multicast
 task read route-map
 task read sonet-sdh
 task read transport
 task read ext-access
 task read filesystem
 task read tty-access
 task read config-mgmt
 task read ip-services
 task read mpls-static
 task read route-policy
 task read host-services
 task read basic-services
 task read config-services
 task read ethernet-services
!
taskgroup READ-WRITE-TGRP
 inherit taskgroup root-lr
 inherit taskgroup cisco-support
!
usergroup READ-ONLY-UGRP
 taskgroup READ-ONLY-TGRP
!
usergroup READ-WRITE-UGRP
 taskgroup READ-ONLY-TGRP
 taskgroup READ-WRITE-TGRP
!
clock timezone UTC UTC
banner motd ^
***************************************************************************
                            NOTICE TO USERS
This is a private computer system and is for authorized use only. Users
(authorized or unauthorized) have no explicit or implicit expectation of
privacy.
Any or all uses of this system and all files on this system may be
intercepted, monitored, recorded, copied, audited, inspected, and disclosed
to authorized site and law enforcement personnel, as well as authorized
officials of other agencies, both domestic and foreign. By using this
system, the user consents to such interception, monitoring, recording,
copying, auditing, inspection, and disclosure at the discretion of the
authorized site or personnel.
Unauthorized or improper use of this system may result in administrative
disciplinary action and civil and criminal penalties. By continuing to
use this system you indicate your awareness of and consent to these terms
and conditions of use. LOG OFF IMMEDIATELY if you do not agree to the
conditions stated in this warning.
$(hostname) vty $(line)
*****************************************************************************
^
logging trap informational
logging events threshold 85
logging events display-location
logging events level informational
logging archive
 device harddisk
 severity informational
 file-size 10
 frequency daily
 archive-size 2047
 archive-length 12
!
logging console disable
logging history informational
logging monitor disable
logging buffered 3000000
logging buffered informational
logging facility local7
'''
            for line in mtso_config[host_data["New Hostname"][:8]]["logging"]:
                conf += f'''
{line}'''
            conf += f'''
logging localfilesize 10000000
logging source-interface MgmtEth0/RP0/CPU0/0 vrf management
logging hostnameprefix {host_data["New Hostname"]}
service timestamps log datetime localtime msec show-timezone
service timestamps debug datetime localtime msec show-timezone
logging events link-status software-interfaces
domain name verizonwireless.com
domain lookup disable
username PAMadmin
 group READ-WRITE-UGRP
 secret 10 $6$bUi.Q0c9.b0k7Q0.$.OFFvdf/DqLWqjvGFOmUKc3V2D6oFzqAT/utUtfy75Ocy1ewvYWQXDw19pnMHaDP2cu3h1z2ICqftFsIOrtlM1
!
username PAMadmingrp
 group READ-WRITE-UGRP
 secret 5 $1$QBVt$sI8r0CpeftNQ7jDA8CaWl/
!
username PAMronlygrp
 group READ-WRITE-UGRP
 secret 5 $1$TXnt$6c/6ENQaZVh3gLupNroUR1
!
username NSOAUTO
 group READ-WRITE-UGRP
 secret 5 $1$smIR$hqLYYlOcbokYbllvL3u5S.
!
username SPOAUTO
 group READ-WRITE-UGRP
 secret 5 $1$g3Ve$Ib89nCw2VnlFGvAQSbpWw1
!
username sev1snmpuser
 group READ-ONLY-UGRP
 secret 5 $1$CEIB$hUvUtmox2sIr6qE0nWLkF0
!
username NCMBBTP
 group READ-WRITE-UGRP
 secret 5 $1$YzvZ$ZEVaeEq4HM23PQbhIxLc1/
!
username NCMSOLK
 group READ-WRITE-UGRP
 secret 5 $1$XbCX$bmq04eJQH3XaPNlgK0wP.1
!
username ienucssnmpusr
 group READ-ONLY-UGRP
 secret 5 $1$pdKk$TRK2/PLE7e2BDrRjCitH8.
!
username PAMvendgrp
 group READ-WRITE-UGRP
 secret 5 $1$F9ah$1VDfl1.b/XFEpsr3FKXFi0
!
username EBHuser
 group READ-WRITE-UGRP
 secret 5 $1$jsq4$NLTN05IFoTI1pR3XeiW.a0
!
username njbbcpnebh
 group READ-WRITE-UGRP
 secret 5 $1$YEL0$V8AvjQSmtM3WxxRUfeC/p0
!
username solkcpnebh
 group READ-ONLY-UGRP
 secret 5 $1$MDA5$NY0xW7ae8RRmY57JjhqGL1
!
aaa authentication login default local
!
'''
            return conf

        def ptp():
            # MOP Section: PTP
            conf = f'''
ptp
 clock
  domain 24
  profile g.8275.1 clock-type T-BC
  timescale PTP
 !
 profile slave
  multicast target-address ethernet 01-1B-19-00-00-00
  transport ethernet
  sync frequency 16
  clock operation one-step
  announce frequency 8
  delay-request frequency 16
 !
 profile master
  multicast target-address ethernet 01-1B-19-00-00-00
  transport ethernet
  sync frequency 16
  announce frequency 8
  delay-request frequency 16
 !
 uncalibrated-clock-class 7
!
'''
            return conf

        def vrfs():
            # MOP Section: VRFs
            conf = f'''
vrf management
 address-family ipv6 unicast
 !
!
vrf RAN
 description VRF 1 - RAN
 address-family ipv6 unicast
  import route-target
   {host_data["COIN ASN"]}:1
  !
  export route-target
   {host_data["COIN ASN"]}:1
  !
 !
!
vrf CELL_MGMT
 description VRF 4 - CELL_MGMT
 address-family ipv4 unicast
  import route-target
   {host_data["COIN ASN"]}:4
  !
  export route-target
   {host_data["COIN ASN"]}:4
  !
 !
 address-family ipv6 unicast
  import route-target
   {host_data["COIN ASN"]}:4
  !
  export route-target
   {host_data["COIN ASN"]}:4
  !
 !
!
'''
            return conf

        def admin_part_2(ntp_srvr1, ntp_srvr2):
            conf = ""
            # MOP Section: Admin Part 2
            conf += f'''
ipv6 path-mtu enable
line console
 length 30
 session-timeout 90
 transport output ssh
!
line default
 session-timeout 30
 transport input ssh
 transport output ssh
!
snmp-server ifmib ifalias long
snmp-server ifindex persist
snmp-server ifmib stats cache
snmp-server trap link ietf
snmp-server mibs cbqosmib persist
snmp-server vrf management'''
            for line in mtso_config[host_data["New Hostname"][:8]]["snmp"]:
                conf += f'''
 {line}'''
            conf += f'''
!
snmp-server user sev1snmpuser sev1group v3 auth md5 encrypted 03274C1D0C1E30787A5C0D341D IPv6 SEV1_ACLv6
snmp-server user ienucssnmpusr snmpV3Pronly v3 auth md5 encrypted 00223F28020B19240B20551C5953 priv des56 encrypted 112F352B1142192E002B32767879 SystemOwner
snmp-server view allmibs system included
snmp-server view allmibs internet included
snmp-server view allmibs interfaces included
snmp-server view allmibs 1.3.6.1 included
snmp-server view allmibs 1.2.840.10006.300 included
snmp-server view allmibs 1.3.6.1.2.1.47 included
snmp-server view allmibs 1.3.6.1.2.1.1.5 included
snmp-server view allmibs 1.3.6.1.2.1.10.166.4.1.3.11 excluded
snmp-server view allmibs 1.3.6.1.4.1.9.9.249.1.1.1.1 excluded
snmp-server community 2Y2LHTZP31 RO IPv6 SNMP_ACLv6
snmp-server community cellbackhaul RW IPv6 SNMP_ACLv6
snmp-server group sev1group v3 auth
snmp-server group snmpV3Pronly v3 auth notify allmibs read allmibs
snmp-server traps rf
snmp-server traps bfd
snmp-server traps ethernet cfm
snmp-server traps ntp
snmp-server traps ethernet oam events
snmp-server traps copy-complete
snmp-server traps snmp
snmp-server traps snmp linkup
snmp-server traps snmp linkdown
snmp-server traps snmp coldstart
snmp-server traps snmp warmstart
snmp-server traps flash removal
snmp-server traps flash insertion
snmp-server traps power
snmp-server traps config
snmp-server traps entity
snmp-server traps selective-vrf-download role-change
snmp-server traps syslog
snmp-server traps system
snmp-server traps optical
snmp-server traps cisco-entity-ext
snmp-server traps entity-state operstatus
snmp-server traps entity-state switchover
snmp-server traps optical-ots
snmp-server traps entity-redundancy all
snmp-server traps entity-redundancy status
snmp-server traps entity-redundancy switchover
snmp-server trap-source MgmtEth0/RP0/CPU0/0
!
ntp'''
            # Record the ntp server IPs from the ODD EBH AL and use them in the config.
            odd_ebh_al = [ciq_db[x]['New Hostname'] for x in ciq_db if 'EBH AL' == ciq_db[x]['eNSESR Role'] and int(ciq_db[x]['New Hostname'][-2:]) % 2 == 1]
            logger.info(f"{host_data['New Hostname']} - NTP config from {odd_ebh_al[0]} will be used.")
            odd_ebh_al_run = CiscoConfParse(os.path.join(os.getcwd(), "NP_DATA", f"{odd_ebh_al[0]}.cfg"), factory=True)
            # Record the ntp server IPs from the existing ODD EBH AL config
            odd_ebh_ntp_srvrs = [x for x in odd_ebh_al_run.find_children_w_parents("^ntp", "server")]
            if odd_ebh_ntp_srvrs:
                for srvr in odd_ebh_ntp_srvrs:
                    srvr_ip = re.search(r"([\da-fA-F]+[.:]+[.:\da-fA-F]+)", srvr)
                    if ":" in srvr_ip.group(1):
                        if "prefer" in srvr:
                            conf += f'''
 server vrf management ipv6 {srvr_ip.group(1)} prefer source MgmtEth0/RP0/CPU0/0'''
                            ntp_srvr1 = srvr_ip.group(1)
                        else:
                            conf += f'''
 server vrf management ipv6 {srvr_ip.group(1)} source MgmtEth0/RP0/CPU0/0'''
                            ntp_srvr2 = srvr_ip.group(1)
            else:
                logger.warning(f"{host_data['New Hostname']} - NTP config from {odd_ebh_al[0]} was not found. Update {odd_ebh_al[0]} and then re-run the script.")

            if not ntp_srvr1 or ntp_srvr2:
                if not ntp_srvr1:
                    logger.warning(f"{host_data['New Hostname']} - Primary NTP server not found in {odd_ebh_al[0]}. Update {odd_ebh_al[0]} and then re-run the script.")
                if not ntp_srvr2:
                    logger.warning(f"{host_data['New Hostname']} - Secondary NTP server not found in {odd_ebh_al[0]}. Update {odd_ebh_al[0]} and then re-run the script.")

            conf += f'''
!
bfd
 multipath include location 0/0/CPU0
!
ipv4 netmask-format bit-count
!
frequency synchronization
!
fpd auto-upgrade enable
hw-module profile qos hqos-enable
!
snmp-server traps isis all
snmp-server traps bgp cbgp2 updown
snmp-server traps bgp updown
!
domain vrf management ipv6 host rcklca63-cisco-license.ntvs.vzwnet.com 2001:4888:a06:1f1b:f1:ff2:0:7
 http client vrf management
 http client source-interface ipv6 MgmtEth0/RP0/CPU0/0
!
call-home
service active
contact-email-addr sch-smart-licensing@cisco.com
profile cisco-sl
  active
  destination address http https://rcklca63-cisco-license.ntvs.vzwnet.com/Transportgateway/services/DeviceRequestHandler
  reporting smart-licensing-data
  destination transport-method http
!'''
            if "55" in host_data["Model"]:
                conf += f'''
license smart flexible-consumption enable'''
            conf += f'''
crypto ca trustpoint Trustpool
crl optional
!
'''
            return conf, ntp_srvr1, ntp_srvr2

        def acls(ntp_srvr1, ntp_srvr2):
            conf = ""
            # MOP Section: ACLs
            conf += f'''
ipv6 access-list MGMT_IN_V6
 5 remark "Version 2023.07.06"
 20 remark "IPv6 Basics"
 21 permit ipv6 fe80::/10 any
 22 permit icmpv6 any any
 90 remark "NTP Server"
 91 permit udp host {ntp_srvr1} any eq ntp
 92 permit udp host {ntp_srvr2} any eq ntp
 100 remark "CyberArk"
 101 permit tcp 2001:4888:a02:2202:a0:fef::/112 any eq ssh
 102 permit tcp 2001:4888:a03:2219:c0:fef::/112 any eq ssh
 103 permit tcp 2001:4888:a02:2208:a0:fef::/112 any eq ssh
 104 permit tcp 2001:4888:a03:221f:c0:fef::/112 any eq ssh
 105 permit tcp 2001:4888:a02:2209:a0:fef::/112 any eq ssh
 106 permit tcp 2001:4888:a03:2220:c0:fef::/112 any eq ssh
 110 remark "NSS DMZ"
 111 permit tcp 2001:4888:a02:1f2b::/64 any eq ssh
 112 permit tcp 2001:4888:a03:1f2b::/64 any eq ssh
 113 permit tcp 2001:4888:a06:1f1b::/64 any eq ssh
 120 remark "iEN UCS"
 121 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq snmp
 122 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq snmp
 123 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq snmp
 124 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq 9980
 125 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq 9980
 126 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq 9980
 127 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq 9990
 128 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq 9990
 129 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq 9990
 130 remark "SyslogNG"
 131 permit udp host 2001:4888:a02:2109:a0:fef:0:84 any eq syslog
 132 permit udp host 2001:4888:a03:2109:c0:fef:0:84 any eq syslog
 140 remark "SevOne Servers and Pollers"
 141 permit udp 2001:4888:a06:1d52:f1:fef::/112 any eq snmp
 142 permit udp 2001:4888:a03:1d12:c0:fef::/112 any eq snmp
 143 permit udp 2001:4888:a02:1d12:a0:fef::/112 any eq snmp
 160 remark "Nokia NSP formerly SAM"
 161 permit udp host 2001:4888:a03:2114:c0:fef:0:2e any eq snmp
 162 permit udp host 2001:4888:a01:2114:a1:fef:0:2e any eq snmp
 170 remark "Syslog ULM"
 171 permit udp 2001:4888:a02:2101:a0:fef::/112 any eq syslog
 172 permit udp 2001:4888:a03:2217:c0:fef::/112 any eq syslog
 173 permit udp 2001:4888:a05:2203:e0:fef::/112 any eq syslog
 174 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq syslog
 180 remark "SANE SSH"
 181 permit tcp 2001:4888:a01:2109:a1:9::/112 any eq ssh
 182 permit tcp 2001:4888:a01:2100:a1:fef::/112 any eq ssh
 183 permit tcp 2001:4888:a03:2144:c0:9::/112 any eq ssh
 184 permit tcp 2001:4888:a03:2100:c0:fef::/112 any eq ssh
 185 permit tcp 2001:4888:a06:2144:f0:9::/112 any eq ssh
 186 permit tcp 2001:4888:a06:2100:f0:fef::/112 any eq ssh
 200 remark "HPNA"
 201 permit tcp host 2001:4888:a02:2104:a0:fef:0:21 any eq ssh
 202 permit tcp host 2001:4888:a03:2110:c0:fef:0:12 any eq ssh
 210 remark "APSN NSOAUTO Service Portal"
 211 permit tcp host 2001:4888:a03:210c:c0:fef:0:20 any eq ssh
 220 remark "Smart Licensing On-Prem SSM"
 221 permit tcp host 2001:4888:a06:1f1b:f1:ff2:0:7 eq https any
 230 remark "EBH-AP"
 231 permit tcp host 2607:f160:8a05:1028:8000::6 any
 240 remark "Cisco CX Cloud"
 241 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc1 any
 242 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc2 any
 243 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc3 any
 244 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc4 any
 400 remark "Labs and FOA Sites Only"
 410 remark "EDN Desktops"
 411 permit tcp 2001:4888:a610::/44 any eq ssh
 412 permit tcp 2001:4888:a620::/44 any eq ssh
 413 permit tcp 2001:4888:a630::/44 any eq ssh
 414 permit tcp 2001:4888:a640::/44 any eq ssh
 415 permit tcp 2001:4888:a650::/44 any eq ssh
 416 permit tcp 2001:4888:a660::/44 any eq ssh
 417 permit tcp 2600:80b:150::/44 any eq ssh
 418 permit tcp 2600:80b:160::/44 any eq ssh
 419 permit tcp 2600:80b:170::/44 any eq ssh
 420 permit tcp 2600:80b:180::/44 any eq ssh
 421 permit tcp 2600:80b:190::/44 any eq ssh
 422 permit tcp 2600:80b:1a0::/44 any eq ssh
 423 permit tcp 2600:80b:1b0::/44 any eq ssh
 424 permit tcp 2600:80b:1c0::/44 any eq ssh
 450 remark "EDN VPN Pools"
 451 permit tcp 2001:4888:a600::/44 any eq ssh
 452 permit tcp 2600:80b:310::/44 any eq ssh
 453 permit tcp 2600:80b:300::/44 any eq ssh
 500 remark VZW Standard SNMP ACL
 501 permit ipv6 2001:4888:a01:2100::/56 any
 502 permit ipv6 2001:4888:a02:2100::/56 any
 503 permit ipv6 2001:4888:a03:2100::/56 any
 504 permit ipv6 2001:4888:a04:2100::/56 any
 505 permit ipv6 2001:4888:a05:2100::/56 any
 506 permit ipv6 2001:4888:a06:2100::/56 any
 507 permit ipv6 2001:4888:a07:2100::/56 any
 508 permit ipv6 2001:4888:a08:2100::/56 any
 509 permit ipv6 2001:4888:a0e:2100::/56 any
 510 permit ipv6 2001:4888:a0f:2100::/56 any
 511 permit ipv6 2001:4888:a06:2200::/56 any
 512 permit ipv6 2001:4888:a02:2200::/56 any
 513 permit ipv6 2001:4888:a02:1d10::/60 any
 514 permit ipv6 2001:4888:a06:1d50::/60 any
 515 permit ipv6 2001:4888:a03:1d10::/60 any
 516 permit ipv6 2001:4888:2:1d10::/60 any
 517 permit ipv6 2001:4888:6:1d50::/60 any
 518 permit ipv6 2001:4888:3:1d10::/60 any
 1000 deny ipv6 any any
!
ipv6 access-list SEV1_ACLv6
 10 remark Version_Q22023
 20 permit ipv6 2001:4888:a02:1d10::/60 any
 30 permit ipv6 2001:4888:a06:1d50::/60 any
 40 permit ipv6 2001:4888:a03:1d10::/60 any
 50 permit ipv6 2001:4888:2:1d10::/60 any
 60 permit ipv6 2001:4888:6:1d50::/60 any
 70 permit ipv6 2001:4888:3:1d10::/60 any
 80 deny ipv6 any any
!
ipv6 access-list SNMP_ACLv6
 10 remark Version_Q22023
 20 permit ipv6 2001:4888:a01:2100::/56 any
 30 permit ipv6 2001:4888:a02:2100::/56 any
 40 permit ipv6 2001:4888:a03:2100::/56 any
 50 permit ipv6 2001:4888:a03:2200::/56 any
 60 permit ipv6 2001:4888:a04:2100::/56 any
 70 permit ipv6 2001:4888:a05:2100::/56 any
 80 permit ipv6 2001:4888:a06:2100::/56 any
 90 permit ipv6 2001:4888:a07:2100::/56 any
 100 permit ipv6 2001:4888:a08:2100::/56 any
 110 permit ipv6 2001:4888:a0e:2100::/56 any
 120 permit ipv6 2001:4888:a0f:2100::/56 any
 130 permit ipv6 2001:4888:a06:2200::/56 any
 140 permit ipv6 2001:4888:a02:2200::/56 any
 150 permit ipv6 2001:4888:A03:2200::/56 any
 160 permit ipv6 2001:4888:a02:1d10::/60 any
 170 permit ipv6 2001:4888:a06:1d50::/60 any
 180 permit ipv6 2001:4888:a03:1d10::/60 any
 190 permit ipv6 2001:4888:2:1d10::/60 any
 200 permit ipv6 2001:4888:6:1d50::/60 any
 210 permit ipv6 2001:4888:3:1d10::/60 any
 220 deny ipv6 any any
!
ipv4 access-list MGMT_IN
 10 remark no IPv4 management access
 20 deny ipv4 any any
!
'''
            for intf in host_data["Interfaces"]:
                if host_data["Interfaces"][intf]["description"].split("_")[0] in ciq_db.keys():
                    if "EBH AL" in ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]["eNSESR Role"]:
                        ebh_al = ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]
            hub_bl = [ciq_db[x] for x in ciq_db if "HUB BL" in ciq_db[x]["eNSESR Role"] and ciq_db[x] is not host_data]
            hub_bl = hub_bl[0]
            if int(ebh_al["New Hostname"][-2:]) % 2 == 0:
                ebh_peer_acl = "EVEN"
            elif int(ebh_al["New Hostname"][-2:]) % 2 == 1:
                ebh_peer_acl = "ODD"
            conf += f'''
ipv4 access-list ACL_BLOCK_EBH_{ebh_peer_acl}
 10 remark Block BFD and BGP from EBH AL {ebh_peer_acl}
 20 deny udp host {ebh_al["Loopback 0 Global - IPv4"]} host {host_data["Loopback 0 Global - IPv4"]} eq 4784
 30 deny udp host {ebh_al["Loopback 0 Global - IPv4"]} eq 4784 host {host_data["Loopback 0 Global - IPv4"]}
 40 deny tcp host {ebh_al["Loopback 0 Global - IPv4"]} host {host_data["Loopback 0 Global - IPv4"]} eq bgp
 50 deny tcp host {ebh_al["Loopback 0 Global - IPv4"]} eq bgp host {host_data["Loopback 0 Global - IPv4"]}
 150 permit ipv4 any any
!
'''
            conf += f'''
ipv4 access-list ACL_BLOCK_BL_PEER
 10 remark Block BFD and BGP from BL PEER
 20 deny udp host {hub_bl["Loopback 0 Global - IPv4"]} host {host_data["Loopback 0 Global - IPv4"]} eq 4784
 30 deny udp host {hub_bl["Loopback 0 Global - IPv4"]} eq 4784 host {host_data["Loopback 0 Global - IPv4"]}
 40 deny tcp host {hub_bl["Loopback 0 Global - IPv4"]} host {host_data["Loopback 0 Global - IPv4"]} eq bgp
 50 deny tcp host {hub_bl["Loopback 0 Global - IPv4"]} eq bgp host {host_data["Loopback 0 Global - IPv4"]}
 150 permit ipv4 any any
!
'''
            conf += f'''
ipv4 access-list ACL_BLOCK_AL_HAIRPIN
 10 remark Block BFD and BGP from EBH AL {ebh_peer_acl}
 20 deny udp host {ebh_al["Loopback 0 Global - IPv4"]} host {host_data["Loopback 0 Global - IPv4"]} eq 4784
 30 deny udp host {ebh_al["Loopback 0 Global - IPv4"]} eq 4784 host {host_data["Loopback 0 Global - IPv4"]}
 40 deny tcp host {ebh_al["Loopback 0 Global - IPv4"]} host {host_data["Loopback 0 Global - IPv4"]} eq bgp
 50 deny tcp host {ebh_al["Loopback 0 Global - IPv4"]} eq bgp host {host_data["Loopback 0 Global - IPv4"]}
 60 remark Block BFD and BGP from BL PEER
 70 deny udp host {hub_bl["Loopback 0 Global - IPv4"]} host {host_data["Loopback 0 Global - IPv4"]} eq 4784
 80 deny udp host {hub_bl["Loopback 0 Global - IPv4"]} eq 4784 host {host_data["Loopback 0 Global - IPv4"]}
 90 deny tcp host {hub_bl["Loopback 0 Global - IPv4"]} host {host_data["Loopback 0 Global - IPv4"]} eq bgp
 100 deny tcp host {hub_bl["Loopback 0 Global - IPv4"]} eq bgp host {host_data["Loopback 0 Global - IPv4"]}
 150 permit ipv4 any any
!
'''
            return conf

        def qos():
            # MOP Section: QOS
            conf = f'''
class-map match-any WPS-IN
 match mpls experimental topmost 7
 match dscp 44
 end-class-map
!
class-map match-any GOLD-IN
 match mpls experimental topmost 3
 match dscp cs3 af21 af22 af23 af31 af32 af33
 end-class-map
!
class-map match-any VOICE-IN
 match mpls experimental topmost 5
 match dscp ef
 end-class-map
!
class-map match-any BRONZE-IN
 match dscp cs1
 match mpls experimental topmost 1
 end-class-map
!
class-map match-any SIGNAL-IN
 match mpls experimental topmost 4
 match dscp cs4 af41 af42 af43
 end-class-map
!
class-map match-any SILVER-IN
 match mpls experimental topmost 2
 match dscp cs2 af11 af12 af13
 end-class-map
!
class-map match-any WPS-QUEUE
 match traffic-class 7
 end-class-map
!
class-map match-any GOLD-QUEUE
 match traffic-class 3
 end-class-map
!
class-map match-any CTRL-BFD-IN
 match mpls experimental topmost 6
 match dscp cs6 cs7
 match precedence 6 7
 end-class-map
!
class-map match-any VOICE-QUEUE
 match traffic-class 5
 end-class-map
!
class-map match-any BRONZE-QUEUE
 match traffic-class 1
 end-class-map
!
class-map match-any SIGNAL-QUEUE
 match traffic-class 4
 end-class-map
!
class-map match-any SILVER-QUEUE
 match traffic-class 2
 end-class-map
!
class-map match-any CTRL-BFD-QUEUE
 match traffic-class 6
 end-class-map
!
policy-map QUEUES-OUT
 class BRONZE-QUEUE
  bandwidth remaining percent 10
  queue-limit 800 ms
 !
 class WPS-QUEUE
  shape average percent 5
  queue-limit 1 ms
  priority level 2
 !
 class CTRL-BFD-QUEUE
  shape average percent 2
  queue-limit 1 ms
  priority level 1
 !
 class VOICE-QUEUE
  shape average percent 60
  priority level 2
  queue-limit 1 ms
 !
 class SIGNAL-QUEUE
  bandwidth remaining percent 15
 !
 class GOLD-QUEUE
  bandwidth remaining percent 10
 !
 class SILVER-QUEUE
  bandwidth remaining percent 10
 !
 class class-default
  bandwidth remaining percent 37
  queue-limit 400 ms
 !
 end-policy-map
!
policy-map CLASSIFY-IN
 class BRONZE-IN
  set traffic-class 1
  set mpls experimental imposition 1
 !
 class WPS-IN
  set traffic-class 7
  set mpls experimental imposition 7
 !
 class CTRL-BFD-IN
  set traffic-class 6
  set mpls experimental imposition 6
 !
 class VOICE-IN
  set mpls experimental imposition 5
  set traffic-class 5
 !
 class SIGNAL-IN
  set mpls experimental imposition 4
  set traffic-class 4
 !
 class GOLD-IN
  set mpls experimental imposition 3
  set traffic-class 3
 !
 class SILVER-IN
  set mpls experimental imposition 2
  set traffic-class 2
 !
 class class-default
  set mpls experimental imposition 0
  set traffic-class 0
 !
 end-policy-map
!
'''
            return conf

        def interfaces_loopback_and_management():
            conf = ""
            # MOP Section: Interfaces - Loopback and Management
            conf += f'''
interface Loopback0
 description Global Loopback
 ipv4 address {host_data["Loopback 0 Global - IPv4"]}/32
!
interface Loopback1
 vrf RAN
 ipv6 address {host_data["Loopback 1 RAN - IPv6"]}/128
!
interface Loopback4
 vrf CELL_MGMT
 ipv6 address {host_data["Loopback 4 CELL_MGMT - IPv6"]}/128
!
interface MgmtEth0/RP0/CPU0/0
 vrf management
 ipv6 address {host_data["MGMT - IPv6"]}/64
 ipv4 access-group MGMT_IN ingress
 ipv6 access-group MGMT_IN_V6 ingress
!
'''
            return conf

        def interfaces_physical(new_conf):
            conf = ""
            # MOP Section: Interfaces - Physical
            # Instantiate the ebh_peer_acl variable.
            ebh_peer_acl = ""
            # Record the EBH AL peer and HUB BL peer for future decisions.
            for intf in host_data["Interfaces"]:
                if "B40" in host_data["Interfaces"][intf]["description"]:
                    ebh_al_peer = ciq_db[host_data["Interfaces"][intf]["description"].split('_')[0]]
                    if int(ebh_al_peer["New Hostname"][-2:]) % 2 == 0:
                        ebh_peer_acl = "EVEN"
                        break
                    elif int(ebh_al_peer["New Hostname"][-2:]) % 2 == 1:
                        ebh_peer_acl = "ODD"
                        break
                    else:
                        logger.critical("HUB BL doesn't have interface connected to EBH AL.")
            # Build new interfaces and shapers as needed.
            for intf in host_data["Interfaces"]:
                if "B40" in host_data["Interfaces"][intf]["description"].split('_')[0]:
                    # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                    intf_sh = intf[:2].lower()
                    if "Breakout" in host_data["Interfaces"][intf].keys():
                        if f'''controller Optics{host_data["Interfaces"][intf]["Breakout"]}''' not in conf or f'''controller Optics{host_data["Interfaces"][intf]["Breakout"]}''' not in new_conf:
                            conf += f'''
controller optics {host_data["Interfaces"][intf]["Breakout"]}
 breakout 4x10
!
'''
                    circuit_rate = int(host_data["Interfaces"][intf]["circuit_rate"])
                    if intf_sh == "gi":
                        if circuit_rate <= 0 or circuit_rate > 1000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 1000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 1000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            intf_bw = shape_rate * 1000
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
  !
 end-policy-map
!
'''
                        elif circuit_rate == 1000:
                            egress_policy = f'''QUEUES-OUT'''
                            intf_bw = 1000000
                    elif intf_sh == "te":
                        if circuit_rate <= 0 or circuit_rate > 10000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 10000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 10000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            intf_bw = shape_rate * 1000
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
  !
 end-policy-map
!
'''
                        elif circuit_rate == 10000:
                            egress_policy = f'''QUEUES-OUT'''
                            intf_bw = 10000000
                    elif intf_sh == "hu":
                        if circuit_rate <= 0 or circuit_rate > 100000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 100000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 100000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            intf_bw = shape_rate * 1000
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
  !
 end-policy-map
!
'''
                        elif circuit_rate == 100000:
                            egress_policy = f'''QUEUES-OUT'''
                            intf_bw = 100000000
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                        exit()
                    conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["cid"]}
 mtu 9100
 load-interval 30
 carrier-delay up 50 down 25
 !
interface {intf}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 bandwidth {intf_bw}
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 ipv4 access-group ACL_BLOCK_BL_PEER ingress
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
 load-interval 30
 encapsulation dot1q {host_data["Interfaces"][intf]["VLAN"]}
!
'''
                elif "B4B" in host_data["Interfaces"][intf]["description"].split('_')[0]:
                    # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                    intf_sh = intf[:2].lower()
                    if "Breakout" in host_data["Interfaces"][intf].keys():
                        if f'''controller optics {host_data["Interfaces"][intf]["Breakout"]}''' not in conf:
                            conf += f'''
controller optics {host_data["Interfaces"][intf]["Breakout"]}
 breakout 4x10
!
'''
                    if intf_sh == "gi":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 1000000
                    elif intf_sh == "te":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 10000000
                    elif intf_sh == "hu":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 100000000
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                        exit()
                    conf += f'''
interface {intf}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 mtu 9100
 load-interval 30
 carrier-delay up 50 down 25
 bandwidth {intf_bw}
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 ipv4 access-group ACL_BLOCK_EBH_{ebh_peer_acl} ingress
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
!
'''
                elif "B4A" in host_data["Interfaces"][intf]["description"].split('_')[0]:
                    # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                    intf_sh = intf[:2].lower()
                    if "Breakout" in host_data["Interfaces"][intf].keys():
                        if f'''controller optics {host_data["Interfaces"][intf]["Breakout"]}''' not in conf:
                            conf += f'''
controller optics {host_data["Interfaces"][intf]["Breakout"]}
 breakout 4x10
!
'''
                    if intf_sh == "gi":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 1000000
                    elif intf_sh == "te":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 10000000
                    elif intf_sh == "hu":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 100000000
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                        exit()
                    conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 mtu 9100
 load-interval 30
 carrier-delay up 50 down 25
 ptp
  profile master
  transport ethernet
  port state master-only
  announce frequency 8
  sync frequency 16
  delay-request frequency 16
 frequency synchronization
 bandwidth {intf_bw}
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 ipv4 access-group ACL_BLOCK_AL_HAIRPIN ingress
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
!
'''
            if "Non-RAN" in host_data.keys():
                # Build new interfaces for any Non-RAN Services.
                for intf in host_data["Non-RAN"]:
                    # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                    intf_sh = intf[:2].lower()
                    if "PTP Southbound" in host_data["Non-RAN"][intf]["Type"]:
                        if "Breakout" in host_data["Non-RAN"][intf].keys():
                            if f'''controller optics {host_data["Non-RAN"][intf]["Breakout"]}''' not in conf and f'''controller optics {host_data["Non-RAN"][intf]["Breakout"]}''' not in new_conf:
                                conf += f'''
controller optics {host_data["Non-RAN"][intf]["Breakout"]}
 breakout 4x10
!
'''
                        if intf_sh == "gi":
                            # If the PTP SB is Gig, include the 'negotiation auto' command.
                            conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]}_{host_data["Non-RAN"][intf]["Non-RAN Device Interface"]}
 ptp
  profile slave
  transport ethernet
  port state slave-only
 !
 negotiation auto
 load-interval 30
 frequency synchronization
  selection input
  wait-to-restore 0
 !
!
interface {intf}.301 l2transport
 description PTP Southbound
 encapsulation dot1q 301
 rewrite ingress tag pop 1 symmetric
!
'''
                        elif intf_sh == "te" or intf_sh == "hu":
                            # If the PTP SB is not Gig, don't include the 'negotiation auto' command.
                            conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]}_{host_data["Non-RAN"][intf]["Non-RAN Device Interface"]}
 ptp
  profile slave
  transport ethernet
  port state slave-only
 !
 load-interval 30
 frequency synchronization
  selection input
  wait-to-restore 0
 !
!
interface {intf}.301 l2transport
 description PTP Southbound
 encapsulation dot1q 301
 rewrite ingress tag pop 1 symmetric
!
'''
                        else:
                            logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                            exit()
            return conf

        def prefix_sets():
            # MOP Section: Prefix-Sets
            conf = f'''
prefix-set PRFX_DEFAULT
  0.0.0.0/0,
  ::/0
end-set
!
'''
            return conf

        def route_policies():
            # MOP Section: Route-Policies
            conf = f'''
route-policy LABEL_LOOPBACK0
  set aigp-metric igp-cost
  set label-index {host_data["ISIS SR Prefix SID"]}
end-policy
!
route-policy SET_ADDPATH
  set path-selection backup 1 install multipath-protect advertise
end-policy
!
route-policy IMPORT_RR-5-ENSESR_BL-EBH_LABEL
 pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_BL-EBH_LABEL
  pass
end-policy
!
route-policy IMPORT_RR-5-ENSESR_BL-EBH_EVPN
 if destination in PRFX_DEFAULT then
  pass
 else
  drop
 endif
end-policy
!
route-policy EXPORT_RR-5-ENSESR_BL-EBH_EVPN
  if destination in PRFX_DEFAULT then
    drop
  elseif evpn-route-type is 5 then
    pass
  else
    drop
  endif
end-policy
!
'''
            return conf

        def static_routes():
            conf = ""
            # MOP Section: Static Routes
            if "EDN-Host" not in host_data:
                logger.critical(f"{host_data['New Hostname']} - EDN Host not listed in this device's dict.")
            else:
                edn_host_bvi = ""
                if "BVI400 - IPv6" in ciq_db[host_data["EDN-Host"]] or "BVI40X - IPv6" in ciq_db[host_data["EDN-Host"]]:
                    if ciq_db[host_data["EDN-Host"]]["BVI400 - IPv6"] is not None and "n/a" not in [x.lower() for x in ciq_db[host_data["EDN-Host"]]["BVI400 - IPv6"]] and "na" not in [x.lower() for x in ciq_db[host_data["EDN-Host"]]["BVI400 - IPv6"]]:
                        edn_host_bvi = ciq_db[host_data["EDN-Host"]]["BVI400 - IPv6"][0]
                    elif ciq_db[host_data["EDN-Host"]]["BVI40X - IPv6"] is not None and "n/a" not in [x.lower() for x in ciq_db[host_data["EDN-Host"]]["BVI40X - IPv6"]] and "na" not in [x.lower() for x in ciq_db[host_data["EDN-Host"]]["BVI40X - IPv6"]]:
                        edn_host_bvi = ciq_db[host_data["EDN-Host"]]["BVI40X - IPv6"][0]
                    else:
                        logger.critical(f'''{ciq_db[host_data["EDN-Host"]]['New Hostname']} - CELL_MGMT BVI is missing.''')
                conf = f'''
router static
 vrf management
  address-family ipv6 unicast
   ::/0 MgmtEth0/RP0/CPU0/0 {edn_host_bvi}
  !
 !
!
'''
            return conf

        def isis():
            conf = ""
            # MOP Section: ISIS Process 5
            # Add base ISIS config.
            conf += f'''
router isis 5
 set-overload-bit on-startup 180
 is-type level-1
 net {host_data["ISIS NET ID"]}
 nsf ietf
 log adjacency changes
 lsp-gen-interval maximum-wait 5000 initial-wait 50 secondary-wait 200
 lsp-refresh-interval 65000
 lsp-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
 max-lsp-lifetime 65535
 address-family ipv4 unicast
  metric-style wide
  microloop avoidance segment-routing
  mpls traffic-eng router-id Loopback0
  spf-interval maximum-wait 5000 initial-wait 50 secondary-wait 200
  segment-routing mpls sr-prefer
  spf prefix-priority critical tag 100
 !
 interface Loopback0
  address-family ipv4 unicast
   prefix-sid index {host_data["ISIS SR Prefix SID"]}
  !
 !'''
            # Add ISIS Interface config.
            for intf in host_data["Interfaces"]:
                if host_data["Interfaces"][intf]["description"].split("_")[0] in ciq_db.keys():
                    if "EBH AL" in ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]["eNSESR Role"]:
                        conf += f'''
 interface {intf}
  circuit-type level-1
  bfd minimum-interval 50
  bfd multiplier 5
  bfd fast-detect ipv4
  point-to-point
  hello-padding disable
  hello-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
  address-family ipv4 unicast
   fast-reroute per-prefix
   fast-reroute per-prefix ti-lfa
   metric 1000000
  !
 !
'''
                    elif "HUB BL" in ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]["eNSESR Role"]:
                        conf += f'''
 interface {intf}
  circuit-type level-1
  bfd minimum-interval 50
  bfd multiplier 5
  bfd fast-detect ipv4
  point-to-point
  hello-padding disable
  hello-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
  address-family ipv4 unicast
   fast-reroute per-prefix
   fast-reroute per-prefix ti-lfa
   metric 300
  !
 !
'''
                    elif "HUB AL" in ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]["eNSESR Role"]:
                        conf += f'''
 interface {intf}
  circuit-type level-1
  bfd minimum-interval 50
  bfd multiplier 5
  bfd fast-detect ipv4
  point-to-point
  hello-padding disable
  hello-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
  address-family ipv4 unicast
   fast-reroute per-prefix
   fast-reroute per-prefix ti-lfa
   metric 300
  !
 !
'''
                else:
                    continue
            return conf

        def bgp():
            conf = ""
            # MOP Section: BGP
            # Add base BGP config.
            conf += f'''
router bgp {host_data["COIN ASN"]}
 bgp router-id {host_data["Loopback 0 Global - IPv4"]}
 bgp graceful-restart
 ibgp policy out enforce-modifications
 !
 address-family ipv4 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
  network {host_data["Loopback 0 Global - IPv4"]}/32 route-policy LABEL_LOOPBACK0
  allocate-label all
 !
 address-family vpnv4 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
 address-family vpnv6 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
 address-family l2vpn evpn
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
 neighbor-group RR-5-ENSESR_EBH
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-policy IMPORT_RR-5-ENSESR_BL-EBH_LABEL in
   route-policy EXPORT_RR-5-ENSESR_BL-EBH_LABEL out
   next-hop-self
  !
  address-family l2vpn evpn
   route-policy IMPORT_RR-5-ENSESR_BL-EBH_EVPN in
   route-policy EXPORT_RR-5-ENSESR_BL-EBH_EVPN out
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
   advertise l2vpn evpn re-originated
  !
 neighbor-group RR-5-ENSESR_AL
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-reflector-client
   next-hop-self
  !
  address-family l2vpn evpn
   route-reflector-client
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
  !
 neighbor-group RR-5-ENSESR_PEER
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-reflector-client
   next-hop-self
  !
  address-family l2vpn evpn
   route-reflector-client
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
   advertise l2vpn evpn re-originated
  !
 !
'''
            # Add BGP Neighbor config.
            for peer in host_data["BGP_Peers"]:
                if "EBH AL" in ciq_db[peer]["eNSESR Role"]:
                    conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_EBH
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                elif "HUB BL" in ciq_db[peer]["eNSESR Role"]:
                    conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_PEER
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                elif "HUB AL" in ciq_db[peer]["eNSESR Role"]:
                    conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_AL
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                else:
                    continue
            # Add VRF config.
            conf += f'''
 vrf RAN
  rd {host_data["Loopback 0 Global - IPv4"]}:1
  bgp router-id {host_data["Loopback 0 Global - IPv4"]}
  address-family ipv6 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
 !
 vrf CELL_MGMT
  rd {host_data["Loopback 0 Global - IPv4"]}:4
  bgp router-id {host_data["Loopback 0 Global - IPv4"]}
  address-family ipv4 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
  address-family ipv6 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
 !
!
'''
            return conf

        def admin_part_4():
            # MOP Section: Admin Part 4
            conf = f'''
snmp-server traps hsrp
snmp-server traps vrrp events
snmp-server traps l2vpn all
snmp-server traps l2vpn cisco
snmp-server traps l2vpn vc-up
snmp-server traps l2vpn vc-down
mpls oam
!
snmp-server traps mpls traffic-eng up
snmp-server traps mpls traffic-eng down
snmp-server traps mpls traffic-eng cisco
snmp-server traps mpls traffic-eng reroute
snmp-server traps mpls traffic-eng cisco-ext preempt
snmp-server traps mpls traffic-eng cisco-ext insuff-bw
snmp-server traps mpls traffic-eng cisco-ext bringup-fail
snmp-server traps mpls traffic-eng cisco-ext reroute-pending
snmp-server traps mpls traffic-eng cisco-ext reroute-pending-clear
snmp-server traps mpls traffic-eng reoptimize
snmp-server traps mpls frr all
snmp-server traps mpls frr protected
snmp-server traps mpls frr unprotected
snmp-server traps mpls ldp up
snmp-server traps mpls ldp down
snmp-server traps mpls ldp threshold
snmp-server traps mpls traffic-eng p2mp up
snmp-server traps mpls traffic-eng p2mp down
snmp-server traps mpls l3vpn all
snmp-server traps mpls l3vpn vrf-up
snmp-server traps mpls l3vpn vrf-down
snmp-server traps mpls l3vpn max-threshold-cleared
snmp-server traps mpls l3vpn max-threshold-exceeded
snmp-server traps mpls l3vpn mid-threshold-exceeded
segment-routing
 global-block 19000 119000
!
snmp-server traps sensor
snmp-server traps fru-ctrl
lldp
 management enable
 extended-show-width enable
!
snmp-server traps l2tun sessions
snmp-server traps l2tun tunnel-up
snmp-server traps l2tun tunnel-down
snmp-server traps l2tun pseudowire status
ssh client source-interface MgmtEth0/RP0/CPU0/0
ssh server enable cipher aes-cbc 3des-cbc
ssh timeout 30
ssh server rate-limit 100
ssh server session-limit 10
ssh server v2
ssh server vrf default
ssh server vrf management
ssh server vrf CELL_MGMT
snmp-server traps fabric plane
snmp-server traps fabric bundle link
snmp-server traps fabric bundle state
end
'''
            return conf

        logger.debug(f"Class: {self.__class__.__name__}")
        conf = ""
        conf += admin_part_1()
        conf += ptp()
        conf += vrfs()
        conf_temp = ""
        ntp_srvr1 = ""
        ntp_srvr2 = ""
        conf_temp, ntp_srvr1, ntp_srvr2 = admin_part_2(ntp_srvr1, ntp_srvr2)
        conf += conf_temp
        conf += acls(ntp_srvr1, ntp_srvr2)
        conf += qos()
        conf += interfaces_loopback_and_management()
        conf += interfaces_physical(conf)
        conf += prefix_sets()
        conf += route_policies()
        conf += static_routes()
        conf += isis()
        conf += bgp()
        conf += admin_part_4()

        # Remove unnecessary blank lines from the new config file.
        # Break up the config into 3 parts. Before banner motd, banner motd, and after banner motd
        conf_parts = conf.split("^", 2)
        conf_p1 = "\n".join([s for s in conf_parts[0].splitlines() if s])
        conf_p2 = conf_parts[1]
        conf_p3 = "\n".join([s for s in conf_parts[2].splitlines() if s])
        conf = "\n".join(["^".join([conf_p1, conf_p2, ""]), conf_p3])
        return conf


class build_hub_al:

    def check(self, now, host_data):

        def pre_check():
            conf = f'''### Pre-Checks ###
end
copy running-config harddisk:PreRunCfgBkp_{now.strftime("%m%d%Y")}.cfg

terminal length 0
show bfd session
show isis adjacency
show isis segment-routing label table
show isis topology
show route vrf all
show route vrf RAN ipv6
show route vrf CELL_MGMT ipv6
show route vrf all ipv6
show arp vrf all
show ipv6 neighbors
show bgp l2vpn evpn summary
show bgp l2vpn evpn
show bgp ipv4 label summary
show bgp ipv4 label
show bgp labels'''
            for intf in host_data['Interfaces']:
                conf += f'''
show interface {intf}
show controller {intf.split(".")[0]}'''
            conf += f'''
show calendar
show ntp status
show ntp associations detail
show policy-map interface all | include "input|output"
show ptp interfaces brief
show ptp advertised-clock
show frequency synchronization summary
'''
            return conf

        def post_check():
            conf = f'''### Post-Checks ###
end
terminal length 0
show bfd session
show isis adjacency
show isis segment-routing label table
show isis topology
show route vrf all
show route vrf RAN ipv6
show route vrf CELL_MGMT ipv6
show route vrf all ipv6
show arp vrf all
show ipv6 neighbors
show bgp l2vpn evpn summary
show bgp l2vpn evpn
show bgp ipv4 label summary
show bgp ipv4 label
show bgp labels'''
            for intf in host_data['Interfaces']:
                conf += f'''
show interface {intf}
show controller {intf.split(".")[0]}'''
            conf += f'''
show calendar
show ntp status
show ntp associations detail
show policy-map interface all | include "input|output"
show ptp interfaces brief
show ptp advertised-clock
show frequency synchronization summary
'''
            return conf

        logger.debug(f"Class: {self.__class__.__name__}")
        conf = pre_check() + "\n" + post_check()
        return conf

    def config(self, host_data, ciq_db, runfile, mtso_config):

        def admin_part_1():
            # MOP Section: Admin Part 1
            conf = f'''
hostname {host_data["New Hostname"]}
taskgroup READ-ONLY-TGRP
 task read fr
 task read li
 task read aaa
 task read acl
 task read atm
 task read bfd
 task read bgp
 task read cdp
 task read cef
 task read cgn
 task read eem
 task read ppp
 task read qos
 task read rib
 task read rip
 task read sbc
 task read ancp
 task read bcdl
 task read boot
 task read diag
 task read dwdm
 task read hdlc
 task read hsrp
 task read ipv4
 task read ipv6
 task read isis
 task read lpts
 task read ospf
 task read ouni
 task read snmp
 task read vlan
 task read vrrp
 task read admin
 task read eigrp
 task read l2vpn
 task read bundle
 task read crypto
 task read fabric
 task read static
 task read sysmgr
 task read system
 task read tunnel
 task read drivers
 task read logging
 task read monitor
 task read mpls-te
 task read netflow
 task read network
 task read pos-dpt
 task read firewall
 task read mpls-ldp
 task read pkg-mgmt
 task read fault-mgr
 task read interface
 task read inventory
 task read multicast
 task read route-map
 task read sonet-sdh
 task read transport
 task read ext-access
 task read filesystem
 task read tty-access
 task read config-mgmt
 task read ip-services
 task read mpls-static
 task read route-policy
 task read host-services
 task read basic-services
 task read config-services
 task read ethernet-services
!
taskgroup READ-WRITE-TGRP
 inherit taskgroup root-lr
 inherit taskgroup cisco-support
!
usergroup READ-ONLY-UGRP
 taskgroup READ-ONLY-TGRP
!
usergroup READ-WRITE-UGRP
 taskgroup READ-ONLY-TGRP
 taskgroup READ-WRITE-TGRP
!
clock timezone UTC UTC
banner motd ^
***************************************************************************
                            NOTICE TO USERS
This is a private computer system and is for authorized use only. Users
(authorized or unauthorized) have no explicit or implicit expectation of
privacy.
Any or all uses of this system and all files on this system may be
intercepted, monitored, recorded, copied, audited, inspected, and disclosed
to authorized site and law enforcement personnel, as well as authorized
officials of other agencies, both domestic and foreign. By using this
system, the user consents to such interception, monitoring, recording,
copying, auditing, inspection, and disclosure at the discretion of the
authorized site or personnel.
Unauthorized or improper use of this system may result in administrative
disciplinary action and civil and criminal penalties. By continuing to
use this system you indicate your awareness of and consent to these terms
and conditions of use. LOG OFF IMMEDIATELY if you do not agree to the
conditions stated in this warning.
$(hostname) vty $(line)
*****************************************************************************
^
logging trap informational
logging events threshold 85
logging events display-location
logging events level informational
logging archive
 device harddisk
 severity informational
 file-size 10
 frequency daily
 archive-size 2047
 archive-length 12
!
logging console disable
logging history informational
logging monitor disable
logging buffered 3000000
logging buffered informational
logging facility local7
'''
            for line in mtso_config[host_data["New Hostname"][:8]]["logging"]:
                conf += f'''
{line}'''
            conf += f'''
logging localfilesize 10000000
logging source-interface MgmtEth0/RP0/CPU0/0 vrf management
logging hostnameprefix {host_data["New Hostname"]}
service timestamps log datetime localtime msec show-timezone
service timestamps debug datetime localtime msec show-timezone
logging events link-status software-interfaces
domain name verizonwireless.com
domain lookup disable
username PAMadmin
 group READ-WRITE-UGRP
 secret 10 $6$bUi.Q0c9.b0k7Q0.$.OFFvdf/DqLWqjvGFOmUKc3V2D6oFzqAT/utUtfy75Ocy1ewvYWQXDw19pnMHaDP2cu3h1z2ICqftFsIOrtlM1
!
username PAMadmingrp
 group READ-WRITE-UGRP
 secret 5 $1$QBVt$sI8r0CpeftNQ7jDA8CaWl/
!
username PAMronlygrp
 group READ-WRITE-UGRP
 secret 5 $1$TXnt$6c/6ENQaZVh3gLupNroUR1
!
username NSOAUTO
 group READ-WRITE-UGRP
 secret 5 $1$smIR$hqLYYlOcbokYbllvL3u5S.
!
username SPOAUTO
 group READ-WRITE-UGRP
 secret 5 $1$g3Ve$Ib89nCw2VnlFGvAQSbpWw1
!
username sev1snmpuser
 group READ-ONLY-UGRP
 secret 5 $1$CEIB$hUvUtmox2sIr6qE0nWLkF0
!
username NCMBBTP
 group READ-WRITE-UGRP
 secret 5 $1$YzvZ$ZEVaeEq4HM23PQbhIxLc1/
!
username NCMSOLK
 group READ-WRITE-UGRP
 secret 5 $1$XbCX$bmq04eJQH3XaPNlgK0wP.1
!
username ienucssnmpusr
 group READ-ONLY-UGRP
 secret 5 $1$pdKk$TRK2/PLE7e2BDrRjCitH8.
!
username PAMvendgrp
 group READ-WRITE-UGRP
 secret 5 $1$F9ah$1VDfl1.b/XFEpsr3FKXFi0
!
username EBHuser
 group READ-WRITE-UGRP
 secret 5 $1$jsq4$NLTN05IFoTI1pR3XeiW.a0
!
username njbbcpnebh
 group READ-WRITE-UGRP
 secret 5 $1$YEL0$V8AvjQSmtM3WxxRUfeC/p0
!
username solkcpnebh
 group READ-ONLY-UGRP
 secret 5 $1$MDA5$NY0xW7ae8RRmY57JjhqGL1
!
aaa authentication login default local
!
'''
            return conf

        def ptp():
            # MOP Section: PTP
            conf = f'''
ptp
 clock
  domain 24
  profile g.8275.1 clock-type T-BC
  timescale PTP
 !
 profile slave
  multicast target-address ethernet 01-1B-19-00-00-00
  transport ethernet
  sync frequency 16
  clock operation one-step
  announce frequency 8
  delay-request frequency 16
 !
 profile master
  multicast target-address ethernet 01-1B-19-00-00-00
  transport ethernet
  sync frequency 16
  announce frequency 8
  delay-request frequency 16
 !
 holdover-spec-duration 1800
 holdover-spec-clock-class 7
 uncalibrated-clock-class 7
 holdover-spec-traceable-override
!
'''
            return conf

        def vrfs():
            # MOP Section: VRFs
            conf = f'''
vrf management
 address-family ipv6 unicast
 !
!
vrf RAN
 description VRF 1 - RAN
 address-family ipv6 unicast
  import route-target
   {host_data["COIN ASN"]}:1
   {host_data["COIN ASN"]}:100 stitching
  !
  export route-target
   {host_data["COIN ASN"]}:1
   {host_data["COIN ASN"]}:100 stitching
  !
 !
!
vrf CELL_MGMT
 description VRF 4 - CELL_MGMT
 address-family ipv4 unicast
  import route-target
   {host_data["COIN ASN"]}:4
   {host_data["COIN ASN"]}:400 stitching
  !
  export route-target
   {host_data["COIN ASN"]}:4
   {host_data["COIN ASN"]}:400 stitching
  !
 !
 address-family ipv6 unicast
  import route-target
   {host_data["COIN ASN"]}:4
   {host_data["COIN ASN"]}:400 stitching
  !
  export route-target
   {host_data["COIN ASN"]}:4
   {host_data["COIN ASN"]}:400 stitching
  !
 !
!
'''
            return conf

        def admin_part_2(ntp_srvr1, ntp_srvr2):
            conf = ""
            # MOP Section: Admin Part 2
            conf += f'''
ipv6 path-mtu enable
line console
 length 30
 session-timeout 90
 transport output ssh
!
line default
 session-timeout 30
 transport input ssh
 transport output ssh
!
snmp-server ifmib ifalias long
snmp-server ifindex persist
snmp-server ifmib stats cache
snmp-server trap link ietf
snmp-server mibs cbqosmib persist
snmp-server vrf management'''
            for line in mtso_config[host_data["New Hostname"][:8]]["snmp"]:
                conf += f'''
 {line}'''
            conf += f'''
!
snmp-server user sev1snmpuser sev1group v3 auth md5 encrypted 03274C1D0C1E30787A5C0D341D IPv6 SEV1_ACLv6
snmp-server user ienucssnmpusr snmpV3Pronly v3 auth md5 encrypted 00223F28020B19240B20551C5953 priv des56 encrypted 112F352B1142192E002B32767879 SystemOwner
snmp-server view allmibs system included
snmp-server view allmibs internet included
snmp-server view allmibs interfaces included
snmp-server view allmibs 1.3.6.1 included
snmp-server view allmibs 1.2.840.10006.300 included
snmp-server view allmibs 1.3.6.1.2.1.47 included
snmp-server view allmibs 1.3.6.1.2.1.1.5 included
snmp-server view allmibs 1.3.6.1.2.1.10.166.4.1.3.11 excluded
snmp-server view allmibs 1.3.6.1.4.1.9.9.249.1.1.1.1 excluded
snmp-server community 2Y2LHTZP31 RO IPv6 SNMP_ACLv6
snmp-server community cellbackhaul RW IPv6 SNMP_ACLv6
snmp-server group sev1group v3 auth
snmp-server group snmpV3Pronly v3 auth notify allmibs read allmibs
snmp-server traps rf
snmp-server traps bfd
snmp-server traps ethernet cfm
snmp-server traps ntp
snmp-server traps ethernet oam events
snmp-server traps copy-complete
snmp-server traps snmp
snmp-server traps snmp linkup
snmp-server traps snmp linkdown
snmp-server traps snmp coldstart
snmp-server traps snmp warmstart
snmp-server traps flash removal
snmp-server traps flash insertion
snmp-server traps power
snmp-server traps config
snmp-server traps entity
snmp-server traps selective-vrf-download role-change
snmp-server traps syslog
snmp-server traps system
snmp-server traps optical
snmp-server traps cisco-entity-ext
snmp-server traps entity-state operstatus
snmp-server traps entity-state switchover
snmp-server traps optical-ots
snmp-server traps entity-redundancy all
snmp-server traps entity-redundancy status
snmp-server traps entity-redundancy switchover
snmp-server trap-source MgmtEth0/RP0/CPU0/0
!
ntp'''
            # Record the ntp server IPs from the ODD EBH AL and use them in the config.
            odd_ebh_al = [ciq_db[x]['New Hostname'] for x in ciq_db if 'EBH AL' == ciq_db[x]['eNSESR Role'] and int(ciq_db[x]['New Hostname'][-2:]) % 2 == 1]
            logger.info(f"{host_data['New Hostname']} - NTP config from {odd_ebh_al[0]} will be used.")
            odd_ebh_al_run = CiscoConfParse(os.path.join(os.getcwd(), "NP_DATA", f"{odd_ebh_al[0]}.cfg"), factory=True)
            # Record the ntp server IPs from the existing ODD EBH AL config
            odd_ebh_ntp_srvrs = [x for x in odd_ebh_al_run.find_children_w_parents("^ntp", "server")]
            if odd_ebh_ntp_srvrs:
                for srvr in odd_ebh_ntp_srvrs:
                    srvr_ip = re.search(r"([\da-fA-F]+[.:]+[.:\da-fA-F]+)", srvr)
                    if ":" in srvr_ip.group(1):
                        if "prefer" in srvr:
                            conf += f'''
 server vrf management ipv6 {srvr_ip.group(1)} prefer source MgmtEth0/RP0/CPU0/0'''
                            ntp_srvr1 = srvr_ip.group(1)
                        else:
                            conf += f'''
 server vrf management ipv6 {srvr_ip.group(1)} source MgmtEth0/RP0/CPU0/0'''
                            ntp_srvr2 = srvr_ip.group(1)

            if not ntp_srvr1 or ntp_srvr2:
                if not ntp_srvr1:
                    logger.warning(f"{host_data['New Hostname']} - Primary NTP server not found in {odd_ebh_al[0]}. Update {odd_ebh_al[0]} and then re-run the script.")
                if not ntp_srvr2:
                    logger.warning(f"{host_data['New Hostname']} - Secondary NTP server not found in {odd_ebh_al[0]}. Update {odd_ebh_al[0]} and then re-run the script.")

            conf += f'''
!
bfd
 multipath include location 0/0/CPU0
!
ipv4 netmask-format bit-count
!
frequency synchronization
!
fpd auto-upgrade enable
hw-module profile qos hqos-enable
!
snmp-server traps isis all
snmp-server traps bgp cbgp2 updown
snmp-server traps bgp updown
!
domain vrf management ipv6 host rcklca63-cisco-license.ntvs.vzwnet.com 2001:4888:a06:1f1b:f1:ff2:0:7
 http client vrf management
 http client source-interface ipv6 MgmtEth0/RP0/CPU0/0
!
call-home
service active
contact-email-addr sch-smart-licensing@cisco.com
profile cisco-sl
  active
  destination address http https://rcklca63-cisco-license.ntvs.vzwnet.com/Transportgateway/services/DeviceRequestHandler
  reporting smart-licensing-data
  destination transport-method http
!'''
            if "55" in host_data["Model"]:
                conf += f'''
license smart flexible-consumption enable'''
            conf += f'''
crypto ca trustpoint Trustpool
crl optional
!
'''
            return conf, ntp_srvr1, ntp_srvr2

        def acls(ntp_srvr1, ntp_srvr2):
            # MOP Section: ACLs
            conf = f'''
ipv6 access-list MGMT_IN_V6
 5 remark "Version 2023.07.06"
 20 remark "IPv6 Basics"
 21 permit ipv6 fe80::/10 any
 22 permit icmpv6 any any
 90 remark "NTP Server"
 91 permit udp host {ntp_srvr1} any eq ntp
 92 permit udp host {ntp_srvr2} any eq ntp
 100 remark "CyberArk"
 101 permit tcp 2001:4888:a02:2202:a0:fef::/112 any eq ssh
 102 permit tcp 2001:4888:a03:2219:c0:fef::/112 any eq ssh
 103 permit tcp 2001:4888:a02:2208:a0:fef::/112 any eq ssh
 104 permit tcp 2001:4888:a03:221f:c0:fef::/112 any eq ssh
 105 permit tcp 2001:4888:a02:2209:a0:fef::/112 any eq ssh
 106 permit tcp 2001:4888:a03:2220:c0:fef::/112 any eq ssh
 110 remark "NSS DMZ"
 111 permit tcp 2001:4888:a02:1f2b::/64 any eq ssh
 112 permit tcp 2001:4888:a03:1f2b::/64 any eq ssh
 113 permit tcp 2001:4888:a06:1f1b::/64 any eq ssh
 120 remark "iEN UCS"
 121 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq snmp
 122 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq snmp
 123 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq snmp
 124 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq 9980
 125 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq 9980
 126 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq 9980
 127 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq 9990
 128 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq 9990
 129 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq 9990
 130 remark "SyslogNG"
 131 permit udp host 2001:4888:a02:2109:a0:fef:0:84 any eq syslog
 132 permit udp host 2001:4888:a03:2109:c0:fef:0:84 any eq syslog
 140 remark "SevOne Servers and Pollers"
 141 permit udp 2001:4888:a06:1d52:f1:fef::/112 any eq snmp
 142 permit udp 2001:4888:a03:1d12:c0:fef::/112 any eq snmp
 143 permit udp 2001:4888:a02:1d12:a0:fef::/112 any eq snmp
 160 remark "Nokia NSP formerly SAM"
 161 permit udp host 2001:4888:a03:2114:c0:fef:0:2e any eq snmp
 162 permit udp host 2001:4888:a01:2114:a1:fef:0:2e any eq snmp
 170 remark "Syslog ULM"
 171 permit udp 2001:4888:a02:2101:a0:fef::/112 any eq syslog
 172 permit udp 2001:4888:a03:2217:c0:fef::/112 any eq syslog
 173 permit udp 2001:4888:a05:2203:e0:fef::/112 any eq syslog
 174 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq syslog
 180 remark "SANE SSH"
 181 permit tcp 2001:4888:a01:2109:a1:9::/112 any eq ssh
 182 permit tcp 2001:4888:a01:2100:a1:fef::/112 any eq ssh
 183 permit tcp 2001:4888:a03:2144:c0:9::/112 any eq ssh
 184 permit tcp 2001:4888:a03:2100:c0:fef::/112 any eq ssh
 185 permit tcp 2001:4888:a06:2144:f0:9::/112 any eq ssh
 186 permit tcp 2001:4888:a06:2100:f0:fef::/112 any eq ssh
 200 remark "HPNA"
 201 permit tcp host 2001:4888:a02:2104:a0:fef:0:21 any eq ssh
 202 permit tcp host 2001:4888:a03:2110:c0:fef:0:12 any eq ssh
 210 remark "APSN NSOAUTO Service Portal"
 211 permit tcp host 2001:4888:a03:210c:c0:fef:0:20 any eq ssh
 220 remark "Smart Licensing On-Prem SSM"
 221 permit tcp host 2001:4888:a06:1f1b:f1:ff2:0:7 eq https any
 230 remark "EBH-AP"
 231 permit tcp host 2607:f160:8a05:1028:8000::6 any
 240 remark "Cisco CX Cloud"
 241 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc1 any
 242 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc2 any
 243 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc3 any
 244 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc4 any
 400 remark "Labs and FOA Sites Only"
 410 remark "EDN Desktops"
 411 permit tcp 2001:4888:a610::/44 any eq ssh
 412 permit tcp 2001:4888:a620::/44 any eq ssh
 413 permit tcp 2001:4888:a630::/44 any eq ssh
 414 permit tcp 2001:4888:a640::/44 any eq ssh
 415 permit tcp 2001:4888:a650::/44 any eq ssh
 416 permit tcp 2001:4888:a660::/44 any eq ssh
 417 permit tcp 2600:80b:150::/44 any eq ssh
 418 permit tcp 2600:80b:160::/44 any eq ssh
 419 permit tcp 2600:80b:170::/44 any eq ssh
 420 permit tcp 2600:80b:180::/44 any eq ssh
 421 permit tcp 2600:80b:190::/44 any eq ssh
 422 permit tcp 2600:80b:1a0::/44 any eq ssh
 423 permit tcp 2600:80b:1b0::/44 any eq ssh
 424 permit tcp 2600:80b:1c0::/44 any eq ssh
 450 remark "EDN VPN Pools"
 451 permit tcp 2001:4888:a600::/44 any eq ssh
 452 permit tcp 2600:80b:310::/44 any eq ssh
 453 permit tcp 2600:80b:300::/44 any eq ssh
 500 remark VZW Standard SNMP ACL
 501 permit ipv6 2001:4888:a01:2100::/56 any
 502 permit ipv6 2001:4888:a02:2100::/56 any
 503 permit ipv6 2001:4888:a03:2100::/56 any
 504 permit ipv6 2001:4888:a04:2100::/56 any
 505 permit ipv6 2001:4888:a05:2100::/56 any
 506 permit ipv6 2001:4888:a06:2100::/56 any
 507 permit ipv6 2001:4888:a07:2100::/56 any
 508 permit ipv6 2001:4888:a08:2100::/56 any
 509 permit ipv6 2001:4888:a0e:2100::/56 any
 510 permit ipv6 2001:4888:a0f:2100::/56 any
 511 permit ipv6 2001:4888:a06:2200::/56 any
 512 permit ipv6 2001:4888:a02:2200::/56 any
 513 permit ipv6 2001:4888:a02:1d10::/60 any
 514 permit ipv6 2001:4888:a06:1d50::/60 any
 515 permit ipv6 2001:4888:a03:1d10::/60 any
 516 permit ipv6 2001:4888:2:1d10::/60 any
 517 permit ipv6 2001:4888:6:1d50::/60 any
 518 permit ipv6 2001:4888:3:1d10::/60 any
 1000 deny ipv6 any any
!
ipv6 access-list SEV1_ACLv6
 10 remark Version_Q22023
 20 permit ipv6 2001:4888:a02:1d10::/60 any
 30 permit ipv6 2001:4888:a06:1d50::/60 any
 40 permit ipv6 2001:4888:a03:1d10::/60 any
 50 permit ipv6 2001:4888:2:1d10::/60 any
 60 permit ipv6 2001:4888:6:1d50::/60 any
 70 permit ipv6 2001:4888:3:1d10::/60 any
 80 deny ipv6 any any
!
ipv6 access-list SNMP_ACLv6
 10 remark Version_Q22023
 20 permit ipv6 2001:4888:a01:2100::/56 any
 30 permit ipv6 2001:4888:a02:2100::/56 any
 40 permit ipv6 2001:4888:a03:2100::/56 any
 50 permit ipv6 2001:4888:a03:2200::/56 any
 60 permit ipv6 2001:4888:a04:2100::/56 any
 70 permit ipv6 2001:4888:a05:2100::/56 any
 80 permit ipv6 2001:4888:a06:2100::/56 any
 90 permit ipv6 2001:4888:a07:2100::/56 any
 100 permit ipv6 2001:4888:a08:2100::/56 any
 110 permit ipv6 2001:4888:a0e:2100::/56 any
 120 permit ipv6 2001:4888:a0f:2100::/56 any
 130 permit ipv6 2001:4888:a06:2200::/56 any
 140 permit ipv6 2001:4888:a02:2200::/56 any
 150 permit ipv6 2001:4888:A03:2200::/56 any
 160 permit ipv6 2001:4888:a02:1d10::/60 any
 170 permit ipv6 2001:4888:a06:1d50::/60 any
 180 permit ipv6 2001:4888:a03:1d10::/60 any
 190 permit ipv6 2001:4888:2:1d10::/60 any
 200 permit ipv6 2001:4888:6:1d50::/60 any
 210 permit ipv6 2001:4888:3:1d10::/60 any
 220 deny ipv6 any any
!
ipv4 access-list MGMT_IN
 10 remark no IPv4 management access
 20 deny ipv4 any any
!
'''
            return conf

        def qos():
            # MOP Section: QOS
            conf = f'''
class-map match-any WPS-IN
 match mpls experimental topmost 7
 match dscp 44
 end-class-map
!
class-map match-any GOLD-IN
 match mpls experimental topmost 3
 match dscp cs3 af21 af22 af23 af31 af32 af33
 end-class-map
!
class-map match-any VOICE-IN
 match mpls experimental topmost 5
 match dscp ef
 end-class-map
!
class-map match-any BRONZE-IN
 match dscp cs1
 match mpls experimental topmost 1
 end-class-map
!
class-map match-any SIGNAL-IN
 match mpls experimental topmost 4
 match dscp cs4 af41 af42 af43
 end-class-map
!
class-map match-any SILVER-IN
 match mpls experimental topmost 2
 match dscp cs2 af11 af12 af13
 end-class-map
!
class-map match-any WPS-QUEUE
 match traffic-class 7
 end-class-map
!
class-map match-any GOLD-QUEUE
 match traffic-class 3
 end-class-map
!
class-map match-any CTRL-BFD-IN
 match mpls experimental topmost 6
 match dscp cs6 cs7
 match precedence 6 7
 end-class-map
!
class-map match-any VOICE-QUEUE
 match traffic-class 5
 end-class-map
!
class-map match-any BRONZE-QUEUE
 match traffic-class 1
 end-class-map
!
class-map match-any SIGNAL-QUEUE
 match traffic-class 4
 end-class-map
!
class-map match-any SILVER-QUEUE
 match traffic-class 2
 end-class-map
!
class-map match-any CTRL-BFD-QUEUE
 match traffic-class 6
 end-class-map
!
policy-map QUEUES-OUT
 class BRONZE-QUEUE
  bandwidth remaining percent 10
  queue-limit 800 ms
 !
 class WPS-QUEUE
  shape average percent 5
  queue-limit 1 ms
  priority level 2
 !
 class CTRL-BFD-QUEUE
  shape average percent 2
  queue-limit 1 ms
  priority level 1
 !
 class VOICE-QUEUE
  shape average percent 60
  priority level 2
  queue-limit 1 ms
 !
 class SIGNAL-QUEUE
  bandwidth remaining percent 15
 !
 class GOLD-QUEUE
  bandwidth remaining percent 10
 !
 class SILVER-QUEUE
  bandwidth remaining percent 10
 !
 class class-default
  bandwidth remaining percent 37
  queue-limit 400 ms
 !
 end-policy-map
!
policy-map CLASSIFY-IN
 class BRONZE-IN
  set traffic-class 1
  set mpls experimental imposition 1
 !
 class WPS-IN
  set traffic-class 7
  set mpls experimental imposition 7
 !
 class CTRL-BFD-IN
  set traffic-class 6
  set mpls experimental imposition 6
 !
 class VOICE-IN
  set mpls experimental imposition 5
  set traffic-class 5
 !
 class SIGNAL-IN
  set mpls experimental imposition 4
  set traffic-class 4
 !
 class GOLD-IN
  set mpls experimental imposition 3
  set traffic-class 3
 !
 class SILVER-IN
  set mpls experimental imposition 2
  set traffic-class 2
 !
 class class-default
  set mpls experimental imposition 0
  set traffic-class 0
 !
 end-policy-map
!
policy-map CLASSIFY-CARRIER-AGG-IN
 class class-default
  set traffic-class 5
  set qos-group 5
 !
 end-policy-map
!
'''
            quads = [x.text for x in runfile.find_objects("hw-module quad")]
            if quads:
                for quad in quads:
                    quad_lines = [x for x in runfile.find_all_children(quad)]
                    for line in quad_lines:
                        conf += f'''{line}\n'''
                    conf += f'''!\n'''
            return conf

        def interfaces_loopback_and_management():
            # MOP Section: Interfaces - Loopback and Management
            conf = f'''
interface Loopback0
 description Global Loopback
 ipv4 address {host_data["Loopback 0 Global - IPv4"]}/32
!
interface Loopback1
 vrf RAN
 ipv6 address {host_data["Loopback 1 RAN - IPv6"]}/128
!
interface Loopback4
 vrf CELL_MGMT
 ipv6 address {host_data["Loopback 4 CELL_MGMT - IPv6"]}/128
!
interface MgmtEth0/RP0/CPU0/0
 vrf management
 ipv6 address {host_data["MGMT - IPv6"]}/64
 ipv4 access-group MGMT_IN ingress
 ipv6 access-group MGMT_IN_V6 ingress
!
'''
            return conf

        def interfaces_physical(prev_conf):
            conf = ""
            # MOP Section: Interfaces - Physical
            # Create a list of interfaces that will be configured based on the CIQ. This will be used when importing legacy interfaces.
            excluded_intf = []
            # Build new interfaces and shapers as needed.
            for intf in host_data["Interfaces"]:
                if "B4B" in host_data["Interfaces"][intf]["description"].split('_')[0]:
                    excluded_intf.append(intf.split('.')[0])
                    # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                    intf_sh = intf[:2].lower()
                    if "Breakout" in host_data["Interfaces"][intf].keys():
                        if f'''controller optics {host_data["Interfaces"][intf]["Breakout"]}''' not in conf:
                            conf += f'''
controller optics {host_data["Interfaces"][intf]["Breakout"]}
 breakout 4x10
!
'''
                    if intf_sh == "gi":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 1000000
                    elif intf_sh == "te":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 10000000
                    elif intf_sh == "hu":
                        egress_policy = f'''QUEUES-OUT'''
                        intf_bw = 100000000
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                        exit()
                    conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 bandwidth {intf_bw}
 mtu 9100
 load-interval 30
 carrier-delay up 50 down 25
 ptp
  profile slave
  transport ethernet
  port state slave-only
  clock operation one-step
  announce frequency 8
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
 frequency synchronization
  selection input
  wait-to-restore 0
!
'''
                if "B4C" in host_data["Interfaces"][intf]["description"].split('_')[0]:
                    excluded_intf.append(intf.split('.')[0])
                    # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                    intf_sh = intf[:2].lower()
                    if "Breakout" in host_data["Interfaces"][intf].keys():
                        if f'''controller Optics{host_data["Interfaces"][intf]["Breakout"]}''' not in conf or f'''controller Optics{host_data["Interfaces"][intf]["Breakout"]}''' not in prev_conf:
                            conf += f'''
controller optics {host_data["Interfaces"][intf]["Breakout"]}
 breakout 4x10
!
'''
                    circuit_rate = int(host_data["Interfaces"][intf]["circuit_rate"])
                    if intf_sh == "gi":
                        if circuit_rate <= 0 or circuit_rate > 1000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 1000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 1000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            intf_bw = shape_rate * 1000
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
  !
 end-policy-map
!
'''
                        elif circuit_rate == 1000:
                            egress_policy = f'''QUEUES-OUT'''
                            intf_bw = 1000000
                    elif intf_sh == "te":
                        if circuit_rate <= 0 or circuit_rate > 10000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 10000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 10000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            intf_bw = shape_rate * 1000
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
  !
 end-policy-map
!
'''
                        elif circuit_rate == 10000:
                            egress_policy = f'''QUEUES-OUT'''
                            intf_bw = 10000000
                    elif intf_sh == "hu":
                        if circuit_rate <= 0 or circuit_rate > 100000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 100000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 100000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            intf_bw = shape_rate * 1000
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
  !
 end-policy-map
!
'''
                        elif circuit_rate == 100000:
                            egress_policy = f'''QUEUES-OUT'''
                            intf_bw = 100000000
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                        exit()
                    conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["cid"]}
 mtu 9100
 load-interval 30
 ptp
  profile master
  transport ethernet
  port state master-only
  clock operation one-step
  announce frequency 8
  sync frequency 16
  delay-request frequency 16
 frequency synchronization
!
interface {intf}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 bandwidth {intf_bw}
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 load-interval 30
 encapsulation dot1q {host_data["Interfaces"][intf]["VLAN"]}
 ipv4 point-to-point
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
!
'''
            # Build new interfaces for any Non-RAN Services as needed.
            if "Non-RAN" in host_data.keys():
                for intf in host_data["Non-RAN"]:
                    excluded_intf.append(intf)
                    # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                    intf_sh = intf[:2].lower()
                    if "PTP Northbound" in host_data["Non-RAN"][intf]["Type"]:
                        if "Breakout" in host_data["Non-RAN"][intf].keys():
                            if f'''controller optics {host_data["Non-RAN"][intf]["Breakout"]}''' not in conf and f'''controller optics {host_data["Non-RAN"][intf]["Breakout"]}''' not in prev_conf:
                                conf += f'''
controller optics {host_data["Non-RAN"][intf]["Breakout"]}
 breakout 4x10
!
'''
                        if intf_sh == "gi":
                            # If the PTP NB is Gig, include the 'negotiation auto' command.
                            conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]}_{host_data["Non-RAN"][intf]["Non-RAN Device Interface"]}
 mtu 1970
 load-interval 30
 negotiation auto
!
interface {intf}.350 l2transport
 no shutdown
 description PTP Northbound
 encapsulation untagged
!
'''
                        elif intf_sh == "te" or intf_sh == "hu":
                            # If the PTP NB is not Gig, don't include the 'negotiation auto' command.
                            conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]}_{host_data["Non-RAN"][intf]["Non-RAN Device Interface"]}
 mtu 1970
 load-interval 30
!
interface {intf}.350 l2transport
 no shutdown
 description PTP Northbound
 encapsulation untagged
!
'''
                        else:
                            logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                            exit()
                    elif "EDN Switch" in host_data["Non-RAN"][intf]["Type"]:
                        if "Breakout" in host_data["Non-RAN"][intf].keys():
                            if f'''controller optics {host_data["Non-RAN"][intf]["Breakout"]}''' not in conf and f'''controller optics {host_data["Non-RAN"][intf]["Breakout"]}''' not in prev_conf:
                                conf += f'''
controller optics {host_data["Non-RAN"][intf]["Breakout"]}
 breakout 4x10
!
'''
                        intf_450 = False
                        if (host_data["BVI450 - IPv4 CIDR"][0] is not None and "n/a" not in host_data["BVI450 - IPv4 CIDR"][0].lower() and "na" not in host_data["BVI450 - IPv4 CIDR"][0].lower()) or (host_data["BVI450 - IPv6"][0] is not None and "n/a" not in host_data["BVI450 - IPv6"][0].lower() and "na" not in host_data["BVI450 - IPv6"][0].lower()):
                            intf_450 = True
                        # Add main interface config.
                        if intf_sh == "gi":
                            # If the EDN SW is Gig, include the 'negotiation auto' command.
                            conf += f'''
interface {intf}
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]}_{host_data["Non-RAN"][intf]["Non-RAN Device Interface"]}
 load-interval 30
 no shutdown
 negotiation auto
!
interface {intf}.400 l2transport
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]}_{host_data["Non-RAN"][intf]["Non-RAN Device Interface"]}
 encapsulation dot1q 400
 rewrite ingress tag pop 1 symmetric
!
'''
                        elif intf_sh == "te" or intf_sh == "hu":
                            # If the EDN SW is not Gig, don't include the 'negotiation auto' command.
                            conf += f'''
interface {intf}
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]}_{host_data["Non-RAN"][intf]["Non-RAN Device Interface"]}
 load-interval 30
 no shutdown
!
interface {intf}.400 l2transport
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]}_{host_data["Non-RAN"][intf]["Non-RAN Device Interface"]}
 encapsulation dot1q 400
 rewrite ingress tag pop 1 symmetric
!
'''
                        else:
                            logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                            exit()
                        # Add sub-interface for vlan 450 if the HUB AL has BVI450 IP assignments.
                        if intf_450:
                            conf += f'''
interface {intf}.450 l2transport
 description vrf CELL_MGMT EDN/IDN UT OAM
 encapsulation dot1q 450
 rewrite ingress tag pop 1 symmetric
!
'''
                    elif "SiteBoss" in host_data["Non-RAN"][intf]["Type"]:
                        if intf_sh == "gi":
                            # If the SiteBoss is Gig, include the 'negotiation auto' command.
                            conf += f'''
interface {intf}
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]} SiteBoss
 service-policy input CLASSIFY-IN
 service-policy output QUEUES-OUT
 load-interval 30
 no shutdown
 negotiation auto
!
interface {intf}.400 l2transport
 description OAM VLAN
 encapsulation untagged
!
'''
                        elif intf_sh == "te" or intf_sh == "hu":
                            # If the SiteBoss is not Gig, don't include the 'negotiation auto' command.
                            conf += f'''
interface {intf}
 description link to {host_data["Non-RAN"][intf]["Non-RAN Device Hostname"]} SiteBoss
 service-policy input CLASSIFY-IN
 service-policy output QUEUES-OUT
 load-interval 30
 no shutdown
!
interface {intf}.400 l2transport
 description OAM VLAN
 encapsulation untagged
!
'''
                        else:
                            logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                            exit()

            # Exclude any Bundles from being carried over.
            bundle_list = [x.text.split(".")[0] for x in runfile.find_objects(r"^interface.+Bundle.+") if "preconfigure" not in x.text]
            excluded_intf += bundle_list
            # Exclude any interfaces to legacy MLS from being carried over.
            backhaul_intf_list = [x.text for x in runfile.find_objects_w_child(r"^interface.+Gig.+", r"^ description.+9010.+") if "preconfigure" not in x.text]
            excluded_intf += backhaul_intf_list
            # Add the backhaul VLAN to the list of excluded interfaces.
            backhaul_vlan_list = [x.split(".")[1] for x in backhaul_intf_list if "." in x]
            excluded_intf += backhaul_vlan_list

            # Create a list of all main interfaces currently configured on the router.
            legacy_intf = [x.text for x in runfile.find_objects(r"^interface.+Gig.+") if "preconfigure" not in x.text]

            # Remove duplicate entries
            legacy_intf_sorted = sorted(set(legacy_intf), key=legacy_intf.index)
            legacy_intf_main_ports = [i for i in legacy_intf_sorted if "." not in i]

            # Create list of interfaces that are currently configured, but aren't listed to be configured in the CIQ.
            excluded_intf_main_ports = [re.findall(r"([\d/]+)", i)[0] for i in excluded_intf if i not in bundle_list]
            excluded_intf_sub_ports = [re.findall(r"([\d/.]+)", i)[0] for i in excluded_intf if i not in bundle_list]
            carried_intf = [l_intf for l_intf in legacy_intf_sorted
                            if re.findall(r"([\d/]+)", l_intf)[0] not in excluded_intf_main_ports
                            and re.findall(r"[\d/]+\.([\d]+).+", l_intf) not in excluded_intf_sub_ports
                            and all(re.findall(r"([\d/]+)", l_intf)[0] != re.findall(r"([\d/]+)", x_intf)[0] for x_intf in excluded_intf_main_ports)
                            if len(re.findall("/", re.findall(r"([\d/]+)", l_intf)[0])) <= 3]
            carried_intf_copy = copy.deepcopy(carried_intf)
            for intf in carried_intf_copy:
                if "." in intf:
                    if re.findall(r"(.+)\.", intf)[0] in legacy_intf_main_ports:
                        continue
                    else:
                        carried_intf.remove(intf)

            # Carry over any currently used shapers.
            for intf in carried_intf:
                try:
                    intf_egress_policy = runfile.find_children_w_parents(intf, "service-policy output")
                except:
                    continue
                else:
                    policies = [x.split(" service-policy output ")[1] for x in intf_egress_policy if "MARK-OUT" not in x]
                    for policy in policies:
                        # If the policy was already added to the config, go to the next policy.
                        if f"policy-map {policy}" in conf or f"policy-map {policy}" in prev_conf:
                            continue
                        elif "-48MB" in policy:
                            if f"policy-map {policy.replace('-48MB', '-NE')}" in conf:
                                continue
                        elif "-46MB" in policy:
                            if f"policy-map {policy.replace('-46MB', '-ST')}" in conf:
                                continue
                        # Get all the config for the egress policy.
                        policy_lines = [x for x in runfile.find_all_children(f"^policy-map {policy}$")]
                        for line in policy_lines:
                            if "-48MB" in line:
                                conf += f'''{line.replace("-48MB", "-NE")}\n'''
                            elif "-46MB" in line:
                                conf += f'''{line.replace("-46MB", "-ST")}\n'''
                            else:
                                conf += f'''{line}\n'''
                        conf += f'''!\n'''

            shut_main_intf_list = [intf for intf in carried_intf if "." not in intf and "shutdown" in [x.strip() for x in runfile.find_all_children(f"^{intf}$")]]

            # Carry over the interface configs.
            for intf in carried_intf:
                shut = False
                if "." in intf:
                    if intf.split(".")[0] in shut_main_intf_list:
                        shut = True
                else:
                    if intf in shut_main_intf_list:
                        shut = True
                intf_lines = [x for x in runfile.find_all_children(f"^{intf}$")]
                for line in intf_lines:
                    # Don't keep the bundle command if it's present.
                    if "bundle" in line.lower():
                        continue
                    # Don't keep the MARK-OUT policy if it's present.
                    elif "mark-out" in line.lower():
                        continue
                    # Change -48MB to -NE
                    elif "-48mb" in line.lower():
                        conf += f'''{line.replace("-48MB", "-NE")}\n'''
                    # Change -46MB to -NE
                    elif "-46mb" in line.lower():
                        conf += f'''{line.replace("-46MB", "-ST")}\n'''
                    elif "interface" in line.lower():
                        conf += f"{line}\n"
                        if not shut:
                            conf += f" no shutdown\n"
                    else:
                        conf += f"{line}\n"
                conf += f"!\n"
            return conf

        def interfaces_bvi():
            conf = ""
            # MOP Section: Interfaces - BVI
            # Create a list of all possible legacy BVIs that were reserved for the HUB AL connections.
            excluded_bvi = [f"interface BVI{x}" for x in range(360, 400, 1)]
            excluded_bvi += [f"interface BVI{x}" for x in range(460, 480, 1)]
            excluded_bvi += ["interface BVI300", "interface BVI400"]
            hub_al_num = int(host_data["New Hostname"][-2:])
            if host_data["BVI10X - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI10X - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI10X - IPv6"]]:
                excluded_bvi.append(f"interface BVI{100 + hub_al_num}")
                if host_data["BVIs need suppress-ra?"] is not None and "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    if len(host_data["BVI10X - IPv6"]) == 1:
                        conf += f'''
interface BVI{100 + hub_al_num}
 description RAN VLAN Interface
 host-routing
 vrf RAN
 ipv6 nd suppress-ra
 ipv6 mtu 1970
 ipv6 address {host_data["BVI10X - IPv6"][0]}/64
 load-interval 30
!
'''
                    elif len(host_data["BVI10X - IPv6"]) > 1:
                        conf += f'''
interface BVI{100 + hub_al_num}
 description RAN VLAN Interface
 host-routing
 vrf RAN
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                        for ipv6_addr in host_data["BVI10X - IPv6"]:
                            conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                elif host_data["BVIs need suppress-ra?"] is not None and "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    if len(host_data["BVI10X - IPv6"]) == 1:
                        conf += f'''
interface BVI{100 + hub_al_num}
 description RAN VLAN Interface
 host-routing
 vrf RAN
 ipv6 mtu 1970
 ipv6 address {host_data["BVI10X - IPv6"][0]}/64
 load-interval 30
!
'''
                    elif len(host_data["BVI10X - IPv6"]) > 1:
                        conf += f'''
interface BVI{100 + hub_al_num}
 description RAN VLAN Interface
 host-routing
 vrf RAN
 ipv6 mtu 1970
'''
                        for ipv6_addr in host_data["BVI10X - IPv6"]:
                            conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            else:
                logger.critical(f"{host_data['New Hostname']} - BVI10X IP not provided in CIQ.")
            if host_data["BVI15X - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI15X - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI15X - IPv6"]]:
                excluded_bvi.append(f"interface BVI{150 + hub_al_num}")
                if host_data["BVIs need suppress-ra?"] is not None and "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    if len(host_data["BVI15X - IPv6"]) == 1:
                        conf += f'''
interface BVI{150 + hub_al_num}
 description RAN VLAN Interface
 host-routing
 vrf RAN
 ipv6 nd suppress-ra
 ipv6 mtu 1970
 ipv6 address {host_data["BVI15X - IPv6"][0]}/64
 load-interval 30
!
'''
                    elif len(host_data["BVI15X - IPv6"]) > 1:
                        conf += f'''
interface BVI{150 + hub_al_num}
 description RAN VLAN Interface
 host-routing
 vrf RAN
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                        for ipv6_addr in host_data["BVI15X - IPv6"]:
                            conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                elif host_data["BVIs need suppress-ra?"] is not None and "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    if len(host_data["BVI15X - IPv6"]) == 1:
                        conf += f'''
interface BVI{150 + hub_al_num}
 description RAN VLAN Interface
 host-routing
 vrf RAN
 ipv6 mtu 1970
 ipv6 address {host_data["BVI15X - IPv6"][0]}/64
 load-interval 30
!
'''
                    elif len(host_data["BVI15X - IPv6"]) > 1:
                        conf += f'''
interface BVI{150 + hub_al_num}
 description RAN VLAN Interface
 host-routing
 vrf RAN
 ipv6 mtu 1970
'''
                        for ipv6_addr in host_data["BVI15X - IPv6"]:
                            conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            else:
                logger.info(f"{host_data['New Hostname']} - BVI15X IP not provided in CIQ.")
            if host_data["BVI40X - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI40X - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI40X - IPv6"]]:
                excluded_bvi.append(f"interface BVI{400 + hub_al_num}")
                if host_data["BVIs need suppress-ra?"] is not None and "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    if len(host_data["BVI40X - IPv6"]) == 1:
                        conf += f'''
interface BVI{400 + hub_al_num}
 description CELL_MGMT VLAN Interface
 host-routing
 vrf CELL_MGMT
 ipv6 nd suppress-ra
 ipv6 mtu 1970
 ipv6 address {host_data["BVI40X - IPv6"][0]}/64
 load-interval 30
!
'''
                    elif len(host_data["BVI40X - IPv6"]) > 1:
                        conf += f'''
interface BVI{400 + hub_al_num}
 description CELL_MGMT VLAN Interface
 host-routing
 vrf CELL_MGMT
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                        for ipv6_addr in host_data["BVI40X - IPv6"]:
                            conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                elif host_data["BVIs need suppress-ra?"] is not None and "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    if len(host_data["BVI40X - IPv6"]) == 1:
                        conf += f'''
interface BVI{400 + hub_al_num}
 description CELL_MGMT VLAN Interface
 host-routing
 vrf CELL_MGMT
 ipv6 mtu 1970
 ipv6 address {host_data["BVI40X - IPv6"][0]}/64
 load-interval 30
!
'''
                    elif len(host_data["BVI40X - IPv6"]) > 1:
                        conf += f'''
interface BVI{400 + hub_al_num}
 description CELL_MGMT VLAN Interface
 host-routing
 vrf CELL_MGMT
 ipv6 mtu 1970
'''
                        for ipv6_addr in host_data["BVI40X - IPv6"]:
                            conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            else:
                logger.critical(f"{host_data['New Hostname']} - BVI40X IP not provided in CIQ.")
            if host_data["BVI350 - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI350 - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI350 - IPv6"]]:
                excluded_bvi.append(f"interface BVI350")
                if host_data["BVIs need suppress-ra?"] is not None and "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    if len(host_data["BVI350 - IPv6"]) == 1:
                        conf += f'''
interface BVI350
 vrf RAN
 ipv6 nd suppress-ra
 ipv6 mtu 1970
 ipv6 address {host_data["BVI350 - IPv6"][0]}/64
 load-interval 30
!
'''
                    elif len(host_data["BVI350 - IPv6"]) > 1:
                        conf += f'''
interface BVI350
 vrf RAN
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                        for ipv6_addr in host_data["BVI350 - IPv6"]:
                            conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                elif host_data["BVIs need suppress-ra?"] is not None and "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    if len(host_data["BVI350 - IPv6"]) == 1:
                        conf += f'''
interface BVI350
 vrf RAN
 ipv6 mtu 1970
 ipv6 address {host_data["BVI350 - IPv6"][0]}/64
 load-interval 30
!
'''
                    elif len(host_data["BVI350 - IPv6"]) > 1:
                        conf += f'''
interface BVI350
 vrf RAN
 ipv6 mtu 1970
'''
                        for ipv6_addr in host_data["BVI350 - IPv6"]:
                            conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            if host_data["BVI400 - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI400 - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI400 - IPv6"]]:
                if "Non-RAN" not in host_data.keys():
                    logger.warning(f"{host_data['New Hostname']} - BVI400 IP provided in CIQ, but this devices is not connected to the EDN switch.")
                else:
                    for intf in host_data["Non-RAN"]:
                        if "EDN Switch" in host_data["Non-RAN"][intf]["Type"]:
                            excluded_bvi.append(f"interface BVI400")
                            if host_data["BVIs need suppress-ra?"] is not None and "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                                if len(host_data["BVI400 - IPv6"]) == 1:
                                    conf += f'''
interface BVI400
 vrf CELL_MGMT
 ipv6 nd suppress-ra
 ipv6 mtu 1970
 ipv6 address {host_data["BVI400 - IPv6"][0]}/64
 load-interval 30
!
'''
                                elif len(host_data["BVI400 - IPv6"]) > 1:
                                    conf += f'''
interface BVI400
 vrf CELL_MGMT
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                                    for ipv6_addr in host_data["BVI400 - IPv6"]:
                                        conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                                    conf += f'''
 load-interval 30
!
'''
                            elif host_data["BVIs need suppress-ra?"] is not None and "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                                if len(host_data["BVI400 - IPv6"]) == 1:
                                    conf += f'''
interface BVI400
 vrf CELL_MGMT
 ipv6 mtu 1970
 ipv6 address {host_data["BVI400 - IPv6"][0]}/64
 load-interval 30
!
'''
                                elif len(host_data["BVI400 - IPv6"]) > 1:
                                    conf += f'''
interface BVI400
 vrf CELL_MGMT
 ipv6 mtu 1970
'''
                                    for ipv6_addr in host_data["BVI400 - IPv6"]:
                                        conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                                    conf += f'''
 load-interval 30
!
'''
                            else:
                                logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            bvi450_ipv6 = False
            bvi450_ipv4 = False
            if host_data["BVI450 - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI450 - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI450 - IPv6"]]:
                bvi450_ipv6 = True
            if host_data["BVI450 - IPv4 CIDR"] is not None and "n/a" not in [x.lower() for x in host_data["BVI450 - IPv4 CIDR"]] and "na" not in [x.lower() for x in host_data["BVI450 - IPv4 CIDR"]]:
                bvi450_ipv4 = True
            if bvi450_ipv6 and bvi450_ipv4:
                excluded_bvi.append(f"interface BVI450")
                if host_data["BVIs need suppress-ra?"] is not None and "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    conf += f'''
interface BVI450
 description vrf CELL_MGMT UT NODES
 host-routing
 vrf CELL_MGMT
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                    for ipv6_addr in host_data["BVI450 - IPv6"]:
                        conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                    for ipv4_addr in host_data["BVI450 - IPv4 CIDR"]:
                        conf += f'''
 ipv4 address {ipv4_addr}
'''
                    conf += f'''
 load-interval 30
!
'''
                elif host_data["BVIs need suppress-ra?"] is not None and "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    conf += f'''
interface BVI450
 description vrf CELL_MGMT UT NODES
 host-routing
 vrf CELL_MGMT
 ipv6 mtu 1970
'''
                    for ipv6_addr in host_data["BVI450 - IPv6"]:
                        conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                    for ipv4_addr in host_data["BVI450 - IPv4 CIDR"]:
                        conf += f'''
 ipv4 address {ipv4_addr}
'''
                    conf += f'''
 load-interval 30
!
'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            elif bvi450_ipv6 and not bvi450_ipv4:
                excluded_bvi.append(f"interface BVI450")
                if host_data["BVIs need suppress-ra?"] is not None and "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    conf += f'''
interface BVI450
 description vrf CELL_MGMT UT NODES
 host-routing
 vrf CELL_MGMT
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                    for ipv6_addr in host_data["BVI450 - IPv6"]:
                        conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                    conf += f'''
 load-interval 30
!
'''
                elif host_data["BVIs need suppress-ra?"] is not None and "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    conf += f'''
interface BVI450
 description vrf CELL_MGMT UT NODES
 host-routing
 vrf CELL_MGMT
 ipv6 mtu 1970
'''
                    for ipv6_addr in host_data["BVI450 - IPv6"]:
                        conf += f'''
 ipv6 address {ipv6_addr}/64
'''
                    conf += f'''
 load-interval 30
!
'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            elif not bvi450_ipv6 and bvi450_ipv4:
                excluded_bvi.append(f"interface BVI450")
                if host_data["BVIs need suppress-ra?"] is not None and "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    conf += f'''
interface BVI450
 description vrf CELL_MGMT UT NODES
 host-routing
 vrf CELL_MGMT
 ipv6 nd suppress-ra
 ipv4 address {host_data["BVI450 - IPv4 CIDR"][0]}
 load-interval 30
!
'''
                elif host_data["BVIs need suppress-ra?"] is not None and "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                    conf += f'''
interface BVI450
 description vrf CELL_MGMT UT NODES
 host-routing
 vrf CELL_MGMT
 ipv4 address {host_data["BVI450 - IPv4 CIDR"][0]}
 load-interval 30
!
'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            # Grab any existing BVIs from the HUB AL
            legacy_bvi = [x.text.split(".")[0] for x in runfile.find_objects(r"^interface.+BVI.+")]
            legacy_bvi_sorted = sorted(set(legacy_bvi), key=legacy_bvi.index)
            carry_bvi = [x for x in legacy_bvi_sorted if x not in excluded_bvi]
            for bvi in carry_bvi:
                bvi_lines = runfile.find_all_children(f"^{bvi}")
                hr_check = [x for x in bvi_lines if "host-routing" in x]
                if hr_check:
                    # Update the BVI if needed.
                    for line in bvi_lines:
                        if "BVI300" in line:
                            bvi_lines[bvi_lines.index(line)] = line.replace("BVI300", "BVI100")
                        if "LTE" in line:
                            bvi_lines[bvi_lines.index(line)] = line.replace("LTE", "RAN")
                    # Add the BVI to the config
                    for line in bvi_lines:
                        conf += f'''{line}\n'''
                    conf += "!\n"
                else:
                    bvi_lines.insert(1, " host-routing")
                    # Update the BVI if needed.
                    for line in bvi_lines:
                        if "BVI300" in line:
                            bvi_lines[bvi_lines.index(line)] = line.replace("BVI300", "BVI100")
                        if "LTE" in line:
                            bvi_lines[bvi_lines.index(line)] = line.replace("LTE", "RAN")
                    # Add the BVI to the config
                    for line in bvi_lines:
                        conf += f'''{line}\n'''
                    conf += "!\n"
            return conf

        def l2vpn(new_conf):
            conf = ""
            # MOP Section: L2VPN
            # Get list of existing bridge domains.
            legacy_bd = [x for x in runfile.find_children_w_parents("^ bridge group", "^  bridge-domain")]
            hub_al_num = int(host_data["New Hostname"][-2:])
            new_bvi10x = 100 + hub_al_num
            new_bvi15x = 150 + hub_al_num
            new_bvi40x = 400 + hub_al_num
            # Instantiate lists that will hold configs
            legacy_bd_300 = []
            legacy_bd_400 = []
            legacy_bd_36x = []
            legacy_bd_38x = []
            legacy_bd_46x = []
            remaining_bd = []
            for bd in legacy_bd:
                # Grab any existing interfaces from the different bridge domains if they exist.
                if "bridge-domain 350" in bd.strip():
                    continue
                elif "bridge-domain 450" in bd.strip():
                    continue
                elif "bridge-domain 300" in bd.strip():
                    legacy_bd_300 = [x for x in runfile.find_children_w_parents("^  bridge-domain 300$", "^   interface") if "bundle" not in x.lower()]
                elif "bridge-domain 400" in bd.strip():
                    legacy_bd_400 = [x for x in runfile.find_children_w_parents("^  bridge-domain 400$", "^   interface") if "bundle" not in x.lower()]
                elif bool(re.search("bridge-domain 3[6-7][0-9]", bd)):
                    legacy_bd_36x = [x for x in runfile.find_children_w_parents("^  bridge-domain 3[6-7][0-9]$", "^   interface") if "bundle" not in x.lower()]
                elif bool(re.search("bridge-domain 3[8-9][0-9]", bd)) and not bool(re.search("bridge-domain 3[6-7][0-9]", bd)):
                    legacy_bd_36x = [x for x in runfile.find_children_w_parents("^  bridge-domain 3[6-9][0-9]$", "^   interface") if "bundle" not in x.lower()]
                elif bool(re.search("bridge-domain 3[8-9][0-9]", bd)):
                    legacy_bd_38x = [x for x in runfile.find_children_w_parents("^  bridge-domain 3[8-9][0-9]$", "^   interface") if "bundle" not in x.lower()]
                elif bool(re.search("bridge-domain 4[6-9][0-9]", bd)):
                    legacy_bd_46x = [x for x in runfile.find_children_w_parents("^  bridge-domain 4[6-9][0-9]$", "^   interface") if "bundle" not in x.lower()]
                else:
                    remaining_bd.append(bd)

            ptp_nb = False
            ptp_sb = False
            edn = False
            edn_intf = ""
            if "Non-RAN" in host_data.keys():
                for intf in host_data["Non-RAN"]:
                    if "PTP Northbound" in host_data["Non-RAN"][intf]["Type"]:
                        ptp_nb = True
                    elif "PTP Southbound" in host_data["Non-RAN"][intf]["Type"]:
                        ptp_sb = True
                    elif "EDN Switch" in host_data["Non-RAN"][intf]["Type"]:
                        edn = True
                        edn_intf = host_data["Non-RAN"][intf]

            bd40x_intf_list = re.findall(r"\ninterface (.+\.400).+l2transport.+\n", new_conf)
            bd450_intf_list = re.findall(r"\ninterface (.+\.450).+l2transport.+\n", new_conf)

            # Add L2VPN config
            conf += f'''
l2vpn
 bridge group VRF_RAN_IRB'''

            # Add config for BD 10x if legacy BD 36x existed.
            if f"interface BVI{new_bvi10x}" in new_conf:
                if len(legacy_bd_36x) > 0:
                    conf += f'''
  bridge-domain {new_bvi10x}'''
                    intf_nums = [re.findall(r"([\d/]+\.[\d]+)", x)[0] for x in legacy_bd_36x if "." in x]
                    carried_intf = [re.search(rf"\n(interface.+{re.escape(x)}).+\n", new_conf)[1] for x in intf_nums if re.search(rf"(\ninterface.+{re.escape(x)}).+\n", new_conf)]
                    # Allow only unique entries to be carried over.
                    carried_intf = sorted(set(carried_intf), key=carried_intf.index)
                    for intf in carried_intf:
                        conf += f'''
   {intf.strip()}
   !'''
                    if len(legacy_bd_300) > 0:
                        intf_nums = [re.findall(r"([\d/]+\.[\d]+)", x)[0] for x in legacy_bd_300 if "." in x]
                        carried_intf = [re.search(rf"\n(interface.+{re.escape(x)}).+\n", new_conf)[1] for x in intf_nums if re.search(rf"(\ninterface.+{re.escape(x)}).+\n", new_conf)]
                        # Allow only unique entries to be carried over.
                        carried_intf = sorted(set(carried_intf), key=carried_intf.index)
                        for intf in carried_intf:
                            conf += f'''
   {intf.strip()}
   !'''
                    conf += f'''
   routed interface BVI{new_bvi10x}
   !
  !'''
                else:
                    conf += f'''
  bridge-domain {new_bvi10x}
   routed interface BVI{new_bvi10x}
   !
  !'''
            else:
                logger.critical(f"{host_data['New Hostname']} - BVI10X wasn't in 'conf', so the L2VPN can't be built.")

            # Add config for BD 15x if legacy BD 38x existed.
            if f"interface BVI{new_bvi15x}" in new_conf:
                if len(legacy_bd_38x) > 0:
                    conf += f'''
  bridge-domain {new_bvi15x}'''
                    intf_nums = [re.findall(r"([\d/]+\.[\d]+)", x)[0] for x in legacy_bd_38x if "." in x]
                    carried_intf = [re.search(rf"\n(interface.+{re.escape(x)}).+\n", new_conf)[1] for x in intf_nums if re.search(rf"(\ninterface.+{re.escape(x)}).+\n", new_conf)]
                    # Allow only unique entries to be carried over.
                    carried_intf = sorted(set(carried_intf), key=carried_intf.index)
                    for intf in carried_intf:
                        conf += f'''
   {intf.strip()}
   !'''
                    conf += f'''
   routed interface BVI{new_bvi15x}
   !
  !'''
            else:
                logger.info(f"{host_data['New Hostname']} - BVI15X wasn't in 'conf', so the L2VPN can't be built.")

            conf += f'''
 bridge group VRF_CELL_MGMT_IRB'''
            # Add config for BD 40x if legacy BD 46x existed.
            if f"interface BVI{new_bvi40x}" in new_conf:
                conf += f'''
  bridge-domain {new_bvi40x}'''
                if len(legacy_bd_46x) > 0:
                    intf_nums = [re.findall(r"([\d/]+\.[\d]+)", x)[0] for x in legacy_bd_46x if "." in x]
                    carried_intf = [re.search(rf"\n(interface.+{re.escape(x)}).+\n", new_conf)[1] for x in intf_nums if re.search(rf"(\ninterface.+{re.escape(x)}).+\n", new_conf)]
                    # Allow only unique entries to be carried over.
                    carried_intf = sorted(set(carried_intf), key=carried_intf.index)
                    for intf in carried_intf:
                        conf += f'''
   {intf.strip()}
   !'''
                    for intf in bd40x_intf_list:
                        conf += f'''
   {intf.strip()}
   !'''

                # Add L2VPN config for EDN OOB if the BVI was configured.
                if edn:
                    if "interface BVI400" not in new_conf:
                        for intf in host_data["Non-RAN"]:
                            if "EDN Switch" in host_data["Non-RAN"][intf]["Type"]:
                                conf += f'''
   interface {intf}.400
   !'''
                conf += f'''
   routed interface BVI{new_bvi40x}
   !
  !
'''
            else:
                logger.critical(f"{host_data['New Hostname']} - BVI40X wasn't in 'conf', so the L2VPN can't be built.")

            # Add L2VPN config for EDN OOB if the BVI was configured.
            if edn:
                if "interface BVI400" in new_conf:
                    conf += f'''
  bridge-domain 400'''
                    for intf in host_data["Non-RAN"]:
                        if "EDN Switch" in host_data["Non-RAN"][intf]["Type"]:
                            conf += f'''
   interface {intf}.400
   !'''
                    conf += f'''
   routed interface BVI400
   !
  !'''
                else:
                    if f"interface BVI{new_bvi40x}" not in new_conf:
                        logger.critical(f"{host_data['New Hostname']} - BVI40X nor BVI400 wasn't in 'conf', so the L2VPN for EDN can't be built.")

            # Add config for BD 450 if BVI450 was created.
            if f"interface BVI450" in new_conf:
                conf += f'''
  bridge-domain 450'''
                if len(bd450_intf_list) > 0:
                    for intf in bd450_intf_list:
                        conf += f'''
   {intf.strip()}
   !'''
                # Add L2VPN config for EDN UT OOB if the BVI was configured.
                if edn:
                    for intf in host_data["Non-RAN"]:
                        if "EDN Switch" in host_data["Non-RAN"][intf]["Type"]:
                            conf += f'''
   interface {intf}.450
   !'''
                conf += f'''
   routed interface BVI450
   !
  !
'''

            # Add any remaining legacy BDs into CELL_MGMT
            if len(remaining_bd) > 0:
                for bd in remaining_bd:
                    legacy_bd_config = [x for x in runfile.find_all_children(f"^{bd}$")]
                    bd_bvi_num = [re.search(r"bvi (\d+)$", x) for x in legacy_bd_config if "bvi" in x]
                    if bd_bvi_num:
                        if f"interface BVI{bd_bvi_num}" in conf:
                            for line in legacy_bd_config:
                                if "bundle" in line.lower():
                                    continue
                                elif "interface" in line:
                                    intf_nums = re.findall(r"([\d/]+\.[\d]+)", line)[0]
                                    carried_intf = ""
                                    if re.search(rf"(\ninterface.+{re.escape(intf_nums)}).+\n", new_conf):
                                        carried_intf = re.search(rf"\n(interface.+{re.escape(intf_nums)}).+\n", new_conf)[1]
                                else:
                                    conf += f'''{line}
    '''
                            conf += f'''  !'''
                        else:
                            logger.critical(f"{host_data['New Hostname']} - BVI{bd_bvi_num} wasn't in 'conf', so the L2VPN can't be built.")
                    else:
                        if len(legacy_bd_config) == 2:
                            # Skip building the legacy BD if there is nothing under it.
                            continue
                        else:
                            for line in legacy_bd_config:
                                if "bundle" in line.lower():
                                    continue
                                else:
                                    conf += f'''{line}
'''
                        conf += f'''  !'''

            # Add L2VPN config for PTP Northbound
            if ptp_nb:
                conf += f'''
 bridge group PTP_NB-BG
  bridge-domain 350'''
                for intf in host_data["Non-RAN"]:
                    if "PTP Northbound" in host_data["Non-RAN"][intf]["Type"]:
                        conf += f'''
   interface {intf}.350
   !'''
                conf += f'''
   routed interface BVI350
   !
  !'''

            conf += f'''
 !
!
'''
            return conf

        def prefix_sets():
            # MOP Section: Prefix-Sets
            # Add a way to gather the EBH GRE prefixes
            conf = f'''
prefix-set PRFX_DEFAULT
  0.0.0.0/0,
  ::/0
end-set
!
prefix-set PRFX_EBH_LOOPBACKS
'''
            if "\n" in host_data["SITE LOOPBACK0 PREFIX"]:
                ebh_prefix_list = [x for x in host_data["SITE LOOPBACK0 PREFIX"].split("\n")]
                for prefix in ebh_prefix_list:
                    conf += f'''
  {prefix} eq 32,'''
            else:
                conf += f'''
  {host_data["SITE LOOPBACK0 PREFIX"]} eq 32,'''
            # Add ODD GRE prefixes.
            if "\n" in host_data["EBH ODD GRE"]:
                odd_gre_list = [x for x in host_data["EBH ODD GRE"].split("\n")]
                for odd_gre in odd_gre_list:
                    conf += f'''
  {odd_gre},'''
            else:
                conf += f'''
  {host_data["EBH ODD GRE"]},'''
            # Add EVEN GRE prefixes.
            if "\n" in host_data["EBH EVEN GRE"]:
                entry_num = 1
                even_gre_list = [x for x in host_data["EBH EVEN GRE"].split("\n")]
                for even_gre in even_gre_list:
                    if entry_num == len(even_gre_list):
                        conf += f'''
  {even_gre}'''
                    else:
                        conf += f'''
  {even_gre},'''
                    entry_num += 1
            else:
                conf += f'''
  {host_data["EBH EVEN GRE"]}'''
            conf += f'''
end-set
!
prefix-set PRFX_GLOBAL_LOOPBACK
  {host_data["Loopback 0 Global - IPv4"]}/32
end-set
!
'''
            return conf

        def route_policies():
            # MOP Section: Route-Policies
            conf = f'''
route-policy LABEL_LOOPBACK0
  set aigp-metric igp-cost
  set label-index {host_data["ISIS SR Prefix SID"]}
end-policy
!
route-policy SET_ADDPATH
  set path-selection backup 1 install multipath-protect advertise
end-policy
!
route-policy IMPORT_RR-5-ENSESR_AL-BL_LABEL
  pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_AL-BL_LABEL
  pass
end-policy
!
route-policy IMPORT_RR-5-ENSESR_AL-CSR_LABEL
  pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_AL-CSR_LABEL
  pass
end-policy
!
route-policy IMPORT_RR-5-ENSESR_AL-CSR_EVPN
  pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_AL-CSR_EVPN
  if destination in PRFX_DEFAULT or evpn-route-type is 5 then
    pass
  else
    drop
  endif
end-policy
!
route-policy IMPORT_RR-5-L3VPN_AL-CSR_LABEL
  pass
end-policy
!
route-policy EXPORT_RR-5-L3VPN_AL-CSR_LABEL
  pass
end-policy
!
route-policy IMPORT_RR-5-L3VPN_AL-CSR_L3VPN
  pass
end-policy
!
route-policy EXPORT_RR-5-L3VPN_AL-CSR_L3VPN
  if destination in PRFX_DEFAULT then
    pass
  endif
end-policy
!
'''
            return conf

        def static_routes():
            conf = ""
            # MOP Section: Static Routes
            if "EDN-Host" not in host_data:
                logger.critical(f"{host_data['New Hostname']} - EDN Host not listed in this device's dict.")
            else:
                edn_host_bvi = ""
                if "BVI400 - IPv6" in ciq_db[host_data["EDN-Host"]] or "BVI40X - IPv6" in ciq_db[host_data["EDN-Host"]]:
                    if ciq_db[host_data["EDN-Host"]]["BVI400 - IPv6"] is not None and "n/a" not in [x.lower() for x in ciq_db[host_data["EDN-Host"]]["BVI400 - IPv6"]] and "na" not in [x.lower() for x in ciq_db[host_data["EDN-Host"]]["BVI400 - IPv6"]]:
                        edn_host_bvi = ciq_db[host_data["EDN-Host"]]["BVI400 - IPv6"][0]
                    elif ciq_db[host_data["EDN-Host"]]["BVI40X - IPv6"] is not None and "n/a" not in [x.lower() for x in ciq_db[host_data["EDN-Host"]]["BVI40X - IPv6"]] and "na" not in [x.lower() for x in ciq_db[host_data["EDN-Host"]]["BVI40X - IPv6"]]:
                        edn_host_bvi = ciq_db[host_data["EDN-Host"]]["BVI40X - IPv6"][0]
                    else:
                        logger.critical(f'''{ciq_db[host_data["EDN-Host"]]['New Hostname']} - CELL_MGMT BVI is missing.''')
                conf = f'''
router static
 vrf management
  address-family ipv6 unicast
   ::/0 MgmtEth0/RP0/CPU0/0 {edn_host_bvi}
  !
 !
!
'''
            return conf

        def isis():
            conf = ""
            # MOP Section: ISIS Process 5
            # Add base ISIS config.
            conf += f'''
router isis 5
 set-overload-bit on-startup 180
 is-type level-1
 net {host_data["ISIS NET ID"]}
 nsf ietf
 log adjacency changes
 lsp-gen-interval maximum-wait 5000 initial-wait 50 secondary-wait 200
 lsp-refresh-interval 65000
 lsp-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
 max-lsp-lifetime 65535
 address-family ipv4 unicast
  metric-style wide
  microloop avoidance segment-routing
  mpls traffic-eng router-id Loopback0
  spf-interval maximum-wait 5000 initial-wait 50 secondary-wait 200
  segment-routing mpls sr-prefer
  spf prefix-priority critical tag 100
 !
 interface Loopback0
  address-family ipv4 unicast
   prefix-sid index {host_data["ISIS SR Prefix SID"]}
  !
 !'''
            # Add ISIS Interface config.
            for intf in host_data["Interfaces"]:
                if host_data["Interfaces"][intf]["description"].split("_")[0] in ciq_db.keys():
                    if "HUB BL" in ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]["eNSESR Role"]:
                        conf += f'''
 interface {intf}
  circuit-type level-1
  bfd minimum-interval 50
  bfd multiplier 5
  bfd fast-detect ipv4
  point-to-point
  hello-padding disable
  hello-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
  address-family ipv4 unicast
   fast-reroute per-prefix
   fast-reroute per-prefix ti-lfa
   metric 300
  !
 !
'''
                    elif "DRAN SPOKE" in ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]["eNSESR Role"]:
                        conf += f'''
 interface {intf}
  circuit-type level-1
  bfd minimum-interval 50
  bfd multiplier 5
  bfd fast-detect ipv4
  point-to-point
  hello-padding disable
  hello-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
  address-family ipv4 unicast
   fast-reroute per-prefix
   fast-reroute per-prefix ti-lfa
   metric 1000000
  !
 !
'''
                else:
                    continue
            return conf

        def bgp():
            conf = ""
            # MOP Section: BGP
            # Add base BGP config.
            conf += f'''
router bgp {host_data["COIN ASN"]}
 bgp router-id {host_data["Loopback 0 Global - IPv4"]}
 bgp graceful-restart
 ibgp policy out enforce-modifications
 !
 address-family ipv4 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
  network {host_data["Loopback 0 Global - IPv4"]}/32 route-policy LABEL_LOOPBACK0
  allocate-label all
 !
 address-family vpnv4 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
 address-family vpnv6 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
 address-family l2vpn evpn
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
 neighbor-group RR-5-ENSESR_CSR
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-policy IMPORT_RR-5-ENSESR_AL-CSR_LABEL in
   route-reflector-client
   route-policy EXPORT_RR-5-ENSESR_AL-CSR_LABEL out
   next-hop-self
  !
  address-family l2vpn evpn
   route-policy IMPORT_RR-5-ENSESR_AL-CSR_EVPN in
   route-reflector-client
   route-policy EXPORT_RR-5-ENSESR_AL-CSR_EVPN out
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
  !
 neighbor-group RR-5-ENSESR_BL
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-policy IMPORT_RR-5-ENSESR_AL-BL_LABEL in
   route-policy EXPORT_RR-5-ENSESR_AL-BL_LABEL out
   next-hop-self
  !
  address-family l2vpn evpn
   import re-originate stitching-rt
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
   advertise l2vpn evpn re-originated
  !
 neighbor-group RR-5-L3VPN_CSR
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-reflector-client
   next-hop-self
   route-policy IMPORT_RR-5-L3VPN_AL-CSR_LABEL in
   route-policy EXPORT_RR-5-L3VPN_AL-CSR_LABEL out
  !
  address-family vpnv4 unicast
   import stitching-rt re-originate
   route-policy IMPORT_RR-5-L3VPN_AL-CSR_L3VPN in
   route-policy EXPORT_RR-5-L3VPN_AL-CSR_L3VPN out
   route-reflector-client
   advertise vpnv4 unicast re-originated stitching-rt
  !
  address-family vpnv6 unicast
   import stitching-rt re-originate
   route-policy IMPORT_RR-5-L3VPN_AL-CSR_L3VPN in
   route-policy EXPORT_RR-5-L3VPN_AL-CSR_L3VPN out
   route-reflector-client
   advertise vpnv6 unicast re-originated stitching-rt
  !
 !
'''
            # Add BGP Neighbor config.
            for peer in host_data["BGP_Peers"]:
                if "HUB BL" in ciq_db[peer]["eNSESR Role"]:
                    conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_BL
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                elif "DRAN SPOKE" in ciq_db[peer]["eNSESR Role"]:
                    if host_data["BGP_Peers"][peer]["Solution"] == "EVPN":
                        conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_CSR
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                    elif host_data["BGP_Peers"][peer]["Solution"] == "L3VPN":
                        conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-L3VPN_CSR
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                else:
                    continue
            # Add VRF config.
            conf += f'''
 vrf RAN
  rd {host_data["Loopback 0 Global - IPv4"]}:1
  bgp router-id {host_data["Loopback 0 Global - IPv4"]}
  address-family ipv6 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
 !
 vrf CELL_MGMT
  rd {host_data["Loopback 0 Global - IPv4"]}:4
  bgp router-id {host_data["Loopback 0 Global - IPv4"]}
  address-family ipv4 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
  address-family ipv6 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
 !
!
'''
            return conf

        def admin_part_4():
            # MOP Section: Admin Part 4
            conf = f'''
snmp-server traps hsrp
snmp-server traps vrrp events
snmp-server traps l2vpn all
snmp-server traps l2vpn cisco
snmp-server traps l2vpn vc-up
snmp-server traps l2vpn vc-down
mpls oam
!
snmp-server traps mpls traffic-eng up
snmp-server traps mpls traffic-eng down
snmp-server traps mpls traffic-eng cisco
snmp-server traps mpls traffic-eng reroute
snmp-server traps mpls traffic-eng cisco-ext preempt
snmp-server traps mpls traffic-eng cisco-ext insuff-bw
snmp-server traps mpls traffic-eng cisco-ext bringup-fail
snmp-server traps mpls traffic-eng cisco-ext reroute-pending
snmp-server traps mpls traffic-eng cisco-ext reroute-pending-clear
snmp-server traps mpls traffic-eng reoptimize
snmp-server traps mpls frr all
snmp-server traps mpls frr protected
snmp-server traps mpls frr unprotected
snmp-server traps mpls ldp up
snmp-server traps mpls ldp down
snmp-server traps mpls ldp threshold
snmp-server traps mpls traffic-eng p2mp up
snmp-server traps mpls traffic-eng p2mp down
snmp-server traps mpls l3vpn all
snmp-server traps mpls l3vpn vrf-up
snmp-server traps mpls l3vpn vrf-down
snmp-server traps mpls l3vpn max-threshold-cleared
snmp-server traps mpls l3vpn max-threshold-exceeded
snmp-server traps mpls l3vpn mid-threshold-exceeded
segment-routing
 global-block 19000 119000
!
snmp-server traps sensor
snmp-server traps fru-ctrl
lldp
 management enable
 extended-show-width enable
!
snmp-server traps l2tun sessions
snmp-server traps l2tun tunnel-up
snmp-server traps l2tun tunnel-down
snmp-server traps l2tun pseudowire status
ssh client source-interface MgmtEth0/RP0/CPU0/0
ssh server enable cipher aes-cbc 3des-cbc
ssh timeout 30
ssh server rate-limit 100
ssh server session-limit 10
ssh server v2
ssh server vrf default
ssh server vrf management
ssh server vrf CELL_MGMT
snmp-server traps fabric plane
snmp-server traps fabric bundle link
snmp-server traps fabric bundle state
end
'''
            return conf

        logger.debug(f"Class: {self.__class__.__name__}")
        conf = ""
        conf += admin_part_1()
        conf += ptp()
        conf += vrfs()
        conf_temp = ""
        ntp_srvr1 = ""
        ntp_srvr2 = ""
        conf_temp, ntp_srvr1, ntp_srvr2 = admin_part_2(ntp_srvr1, ntp_srvr2)
        conf += conf_temp
        conf += acls(ntp_srvr1, ntp_srvr2)
        conf += qos()
        conf += interfaces_loopback_and_management()
        conf += interfaces_physical(conf)
        conf += interfaces_bvi()
        conf += l2vpn(conf)
        conf += prefix_sets()
        conf += route_policies()
        conf += static_routes()
        conf += isis()
        conf += bgp()
        conf += admin_part_4()

        # Remove unnecessary blank lines from the new config file.
        # Break up the config into 3 parts. Before banner motd, banner motd, and after banner motd
        conf_parts = conf.split("^", 2)
        conf_p1 = "\n".join([s for s in conf_parts[0].splitlines() if s])
        conf_p2 = conf_parts[1]
        conf_p3 = "\n".join([s for s in conf_parts[2].splitlines() if s])
        conf = "\n".join(["^".join([conf_p1, conf_p2, ""]), conf_p3])
        return conf


class build_dran_spoke:

    def check(self, now, host_data):

        def pre_check():
            conf = f'''### Pre-Checks ###
end
copy running-config harddisk:PreRunCfgBkp_{now.strftime("%m%d%Y")}.cfg

terminal length 0
show bfd session
show isis adjacency
show isis segment-routing label table
show isis topology
show route vrf all
show route vrf RAN ipv6
show route vrf CELL_MGMT ipv6
show route vrf all ipv6
show arp vrf all
show ipv6 neighbors
show bgp l2vpn evpn summary
show bgp l2vpn evpn
show bgp ipv4 label summary
show bgp ipv4 label
show bgp labels'''
            for intf in host_data['Interfaces']:
                conf += f'''
show interface {intf}
show controller {intf.split(".")[0]}'''
            conf += f'''
show calendar
show ntp status
show ntp associations detail
show policy-map interface all | include "input|output"
show ptp interfaces brief
show ptp advertised-clock
show frequency synchronization summary
'''
            return conf

        def post_check():
            conf = f'''### Post-Checks ###
end
terminal length 0
show bfd session
show isis adjacency
show isis segment-routing label table
show isis topology
show route vrf all
show route vrf RAN ipv6
show route vrf CELL_MGMT ipv6
show route vrf all ipv6
show arp vrf all
show ipv6 neighbors
show bgp l2vpn evpn summary
show bgp l2vpn evpn
show bgp ipv4 label summary
show bgp ipv4 label
show bgp labels'''
            for intf in host_data['Interfaces']:
                conf += f'''
show interface {intf}
show controller {intf.split(".")[0]}'''
            conf += f'''
show calendar
show ntp status
show ntp associations detail
show policy-map interface all | include "input|output"
show ptp interfaces brief
show ptp advertised-clock
show frequency synchronization summary
'''
            return conf

        logger.debug(f"Class: {self.__class__.__name__}")
        conf = pre_check() + "\n" + post_check()
        return conf

    def config(self, host_data, ciq_db, runfile, mtso_config):

        def admin_part_1():
            # MOP Section: Admin Part 1
            conf = f'''
hostname {host_data["New Hostname"]}
taskgroup READ-ONLY-TGRP
 task read fr
 task read li
 task read aaa
 task read acl
 task read atm
 task read bfd
 task read bgp
 task read cdp
 task read cef
 task read cgn
 task read eem
 task read ppp
 task read qos
 task read rib
 task read rip
 task read sbc
 task read ancp
 task read bcdl
 task read boot
 task read diag
 task read dwdm
 task read hdlc
 task read hsrp
 task read ipv4
 task read ipv6
 task read isis
 task read lpts
 task read ospf
 task read ouni
 task read snmp
 task read vlan
 task read vrrp
 task read admin
 task read eigrp
 task read l2vpn
 task read bundle
 task read crypto
 task read fabric
 task read static
 task read sysmgr
 task read system
 task read tunnel
 task read drivers
 task read logging
 task read monitor
 task read mpls-te
 task read netflow
 task read network
 task read pos-dpt
 task read firewall
 task read mpls-ldp
 task read pkg-mgmt
 task read fault-mgr
 task read interface
 task read inventory
 task read multicast
 task read route-map
 task read sonet-sdh
 task read transport
 task read ext-access
 task read filesystem
 task read tty-access
 task read config-mgmt
 task read ip-services
 task read mpls-static
 task read route-policy
 task read host-services
 task read basic-services
 task read config-services
 task read ethernet-services
!
taskgroup READ-WRITE-TGRP
 inherit taskgroup root-lr
 inherit taskgroup cisco-support
!
usergroup READ-ONLY-UGRP
 taskgroup READ-ONLY-TGRP
!
usergroup READ-WRITE-UGRP
 taskgroup READ-ONLY-TGRP
 taskgroup READ-WRITE-TGRP
!
clock timezone UTC UTC
banner motd ^
***************************************************************************
                            NOTICE TO USERS
This is a private computer system and is for authorized use only. Users
(authorized or unauthorized) have no explicit or implicit expectation of
privacy.
Any or all uses of this system and all files on this system may be
intercepted, monitored, recorded, copied, audited, inspected, and disclosed
to authorized site and law enforcement personnel, as well as authorized
officials of other agencies, both domestic and foreign. By using this
system, the user consents to such interception, monitoring, recording,
copying, auditing, inspection, and disclosure at the discretion of the
authorized site or personnel.
Unauthorized or improper use of this system may result in administrative
disciplinary action and civil and criminal penalties. By continuing to
use this system you indicate your awareness of and consent to these terms
and conditions of use. LOG OFF IMMEDIATELY if you do not agree to the
conditions stated in this warning.
$(hostname) vty $(line)
*****************************************************************************
^
logging trap informational
logging events threshold 85
logging events display-location
logging events level informational
logging archive
 device harddisk
 severity informational
 file-size 10
 frequency daily
 archive-size 2047
 archive-length 12
!
logging console disable
logging history informational
logging monitor disable
logging buffered 3000000
logging buffered informational
logging facility local7
'''
            for line in mtso_config[host_data["New Hostname"][:8]]["logging"]:
                conf += f'''
{line}'''
            conf += f'''
logging localfilesize 10000000
logging source-interface MgmtEth0/RP0/CPU0/0 vrf management
logging hostnameprefix {host_data["New Hostname"]}
service timestamps log datetime localtime msec show-timezone
service timestamps debug datetime localtime msec show-timezone
logging events link-status software-interfaces
domain name verizonwireless.com
domain lookup disable
username PAMadmin
 group READ-WRITE-UGRP
 secret 10 $6$bUi.Q0c9.b0k7Q0.$.OFFvdf/DqLWqjvGFOmUKc3V2D6oFzqAT/utUtfy75Ocy1ewvYWQXDw19pnMHaDP2cu3h1z2ICqftFsIOrtlM1
!
username PAMadmingrp
 group READ-WRITE-UGRP
 secret 5 $1$QBVt$sI8r0CpeftNQ7jDA8CaWl/
!
username PAMronlygrp
 group READ-WRITE-UGRP
 secret 5 $1$TXnt$6c/6ENQaZVh3gLupNroUR1
!
username NSOAUTO
 group READ-WRITE-UGRP
 secret 5 $1$smIR$hqLYYlOcbokYbllvL3u5S.
!
username SPOAUTO
 group READ-WRITE-UGRP
 secret 5 $1$g3Ve$Ib89nCw2VnlFGvAQSbpWw1
!
username sev1snmpuser
 group READ-ONLY-UGRP
 secret 5 $1$CEIB$hUvUtmox2sIr6qE0nWLkF0
!
username NCMBBTP
 group READ-WRITE-UGRP
 secret 5 $1$YzvZ$ZEVaeEq4HM23PQbhIxLc1/
!
username NCMSOLK
 group READ-WRITE-UGRP
 secret 5 $1$XbCX$bmq04eJQH3XaPNlgK0wP.1
!
username ienucssnmpusr
 group READ-ONLY-UGRP
 secret 5 $1$pdKk$TRK2/PLE7e2BDrRjCitH8.
!
username PAMvendgrp
 group READ-WRITE-UGRP
 secret 5 $1$F9ah$1VDfl1.b/XFEpsr3FKXFi0
!
username EBHuser
 group READ-WRITE-UGRP
 secret 5 $1$jsq4$NLTN05IFoTI1pR3XeiW.a0
!
username njbbcpnebh
 group READ-WRITE-UGRP
 secret 5 $1$YEL0$V8AvjQSmtM3WxxRUfeC/p0
!
username solkcpnebh
 group READ-ONLY-UGRP
 secret 5 $1$MDA5$NY0xW7ae8RRmY57JjhqGL1
!
aaa authentication login default local
!
'''
            return conf

        def ptp():
            # MOP Section: PTP
            conf = ""
            gnss_lines = [x for x in runfile.find_all_children(r"^gnss-receiver")]
            if gnss_lines:
                gnss_lines.insert(1, " no shutdown")
                for line in gnss_lines:
                    conf += f"{line}\n"
                conf += f'''
 !
!
ptp
 clock
  domain 24
  profile g.8275.1 clock-type T-GM
  timescale PTP
 !
 profile master
  multicast target-address ethernet 01-1B-19-00-00-00
  transport ethernet
  sync frequency 16
  announce frequency 8
  delay-request frequency 16
!
 holdover-spec-duration 1800
 holdover-spec-clock-class 7
 uncalibrated-clock-class 7
 holdover-spec-traceable-override
!
'''
            else:
                conf += f'''
ptp
 clock
  domain 24
  profile g.8275.1 clock-type T-BC
  timescale PTP
 !
 profile slave
  multicast target-address ethernet 01-1B-19-00-00-00
  transport ethernet
  sync frequency 16
  clock operation one-step
  announce frequency 8
  delay-request frequency 16
 !
 profile master
  multicast target-address ethernet 01-1B-19-00-00-00
  transport ethernet
  sync frequency 16
  announce frequency 8
  delay-request frequency 16
 !
 holdover-spec-duration 1800
 holdover-spec-clock-class 7
 uncalibrated-clock-class 7
 holdover-spec-traceable-override
!
'''
            return conf

        def vrfs():
            # MOP Section: VRFs
            conf = f'''
vrf management
 address-family ipv6 unicast
 !
!
vrf RAN
 description VRF 1 - RAN
 address-family ipv6 unicast
  import route-target
   {host_data["COIN ASN"]}:1
  !
  export route-target
   {host_data["COIN ASN"]}:1
  !
 !
!
vrf CELL_MGMT
 description VRF 4 - CELL_MGMT
 address-family ipv4 unicast
  import route-target
   {host_data["COIN ASN"]}:4
  !
  export route-target
   {host_data["COIN ASN"]}:4
  !
 !
 address-family ipv6 unicast
  import route-target
   {host_data["COIN ASN"]}:4
  !
  export route-target
   {host_data["COIN ASN"]}:4
  !
 !
!
'''
            return conf

        def admin_part_2(ntp_srvr1, ntp_srvr2):
            conf = ""
            # MOP Section: Admin Part 2
            conf += f'''
ipv6 path-mtu enable
line console
 length 30
 session-timeout 90
 transport output ssh
!
line default
 session-timeout 30
 transport input ssh
 transport output ssh
!
snmp-server ifmib ifalias long
snmp-server ifindex persist
snmp-server ifmib stats cache
snmp-server trap link ietf
snmp-server mibs cbqosmib persist
snmp-server vrf management'''
            for line in mtso_config[host_data["New Hostname"][:8]]["snmp"]:
                conf += f'''
 {line}'''
            conf += f'''
!
snmp-server user sev1snmpuser sev1group v3 auth md5 encrypted 03274C1D0C1E30787A5C0D341D IPv6 SEV1_ACLv6
snmp-server user ienucssnmpusr snmpV3Pronly v3 auth md5 encrypted 00223F28020B19240B20551C5953 priv des56 encrypted 112F352B1142192E002B32767879 SystemOwner
snmp-server view allmibs system included
snmp-server view allmibs internet included
snmp-server view allmibs interfaces included
snmp-server view allmibs 1.3.6.1 included
snmp-server view allmibs 1.2.840.10006.300 included
snmp-server view allmibs 1.3.6.1.2.1.47 included
snmp-server view allmibs 1.3.6.1.2.1.1.5 included
snmp-server view allmibs 1.3.6.1.2.1.10.166.4.1.3.11 excluded
snmp-server view allmibs 1.3.6.1.4.1.9.9.249.1.1.1.1 excluded
snmp-server community 2Y2LHTZP31 RO IPv6 SNMP_ACLv6
snmp-server community cellbackhaul RW IPv6 SNMP_ACLv6
snmp-server group sev1group v3 auth
snmp-server group snmpV3Pronly v3 auth notify allmibs read allmibs
snmp-server traps rf
snmp-server traps bfd
snmp-server traps ethernet cfm
snmp-server traps ntp
snmp-server traps ethernet oam events
snmp-server traps copy-complete
snmp-server traps snmp
snmp-server traps snmp linkup
snmp-server traps snmp linkdown
snmp-server traps snmp coldstart
snmp-server traps snmp warmstart
snmp-server traps flash removal
snmp-server traps flash insertion
snmp-server traps power
snmp-server traps config
snmp-server traps entity
snmp-server traps selective-vrf-download role-change
snmp-server traps syslog
snmp-server traps system
snmp-server traps optical
snmp-server traps cisco-entity-ext
snmp-server traps entity-state operstatus
snmp-server traps entity-state switchover
snmp-server traps optical-ots
snmp-server traps entity-redundancy all
snmp-server traps entity-redundancy status
snmp-server traps entity-redundancy switchover
snmp-server trap-source MgmtEth0/RP0/CPU0/0
fpd auto-upgrade enable
ntp'''
            # Record the ntp server IPs from the existing ODD EBH AL config
            odd_ebh_al = [ciq_db[x]['New Hostname'] for x in ciq_db if 'EBH AL' == ciq_db[x]['eNSESR Role'] and int(ciq_db[x]['New Hostname'][-2:]) % 2 == 1]
            logger.info(f"{host_data['New Hostname']} - NTP config from {odd_ebh_al[0]} will be used.")
            odd_ebh_al_run = CiscoConfParse(os.path.join(os.getcwd(), "NP_DATA", f"{odd_ebh_al[0]}.cfg"), factory=True)
            # Record the ntp server IPs from the existing ODD EBH AL config
            odd_ebh_ntp_srvrs = [x for x in odd_ebh_al_run.find_children_w_parents("^ntp", "server")]
            if odd_ebh_ntp_srvrs:
                for srvr in odd_ebh_ntp_srvrs:
                    srvr_ip = re.search(r"([\da-fA-F]+[.:]+[.:\da-fA-F]+)", srvr)
                    if ":" in srvr_ip.group(1):
                        if "prefer" in srvr:
                            conf += f'''
 server vrf management ipv6 {srvr_ip.group(1)} prefer source MgmtEth0/RP0/CPU0/0'''
                            ntp_srvr1 = srvr_ip.group(1)
                        else:
                            conf += f'''
 server vrf management ipv6 {srvr_ip.group(1)} source MgmtEth0/RP0/CPU0/0'''
                            ntp_srvr2 = srvr_ip.group(1)

            if not ntp_srvr1 or ntp_srvr2:
                if not ntp_srvr1:
                    logger.warning(f"{host_data['New Hostname']} - Primary NTP server not found in {odd_ebh_al[0]}. Update {odd_ebh_al[0]} and then re-run the script.")
                if not ntp_srvr2:
                    logger.warning(f"{host_data['New Hostname']} - Secondary NTP server not found in {odd_ebh_al[0]}. Update {odd_ebh_al[0]} and then re-run the script.")

            conf += f'''
!
bfd
 multipath include location 0/0/CPU0
!
ipv4 netmask-format bit-count
!
domain vrf management ipv6 host rcklca63-cisco-license.ntvs.vzwnet.com 2001:4888:a06:1f1b:f1:ff2:0:7
 http client vrf management
 http client source-interface ipv6 MgmtEth0/RP0/CPU0/0
!
frequency synchronization
!
call-home
service active
contact-email-addr sch-smart-licensing@cisco.com
profile cisco-sl
  active
  destination address http https://rcklca63-cisco-license.ntvs.vzwnet.com/Transportgateway/services/DeviceRequestHandler
  reporting smart-licensing-data
  destination transport-method http
!'''
            if "55" in host_data["Model"]:
                conf += f'''
license smart flexible-consumption enable'''
            conf += f'''
crypto ca trustpoint Trustpool
crl optional
!
hw-module profile qos hqos-enable
'''
            return conf, ntp_srvr1, ntp_srvr2

        def acls(ntp_srvr1, ntp_srvr2):
            # MOP Section: ACLs
            conf = f'''
ipv6 access-list MGMT_IN_V6
 5 remark "Version 2023.07.06"
 20 remark "IPv6 Basics"
 21 permit ipv6 fe80::/10 any
 22 permit icmpv6 any any
 90 remark "NTP Server"
 91 permit udp host {ntp_srvr1} any eq ntp
 92 permit udp host {ntp_srvr2} any eq ntp
 100 remark "CyberArk"
 101 permit tcp 2001:4888:a02:2202:a0:fef::/112 any eq ssh
 102 permit tcp 2001:4888:a03:2219:c0:fef::/112 any eq ssh
 103 permit tcp 2001:4888:a02:2208:a0:fef::/112 any eq ssh
 104 permit tcp 2001:4888:a03:221f:c0:fef::/112 any eq ssh
 105 permit tcp 2001:4888:a02:2209:a0:fef::/112 any eq ssh
 106 permit tcp 2001:4888:a03:2220:c0:fef::/112 any eq ssh
 110 remark "NSS DMZ"
 111 permit tcp 2001:4888:a02:1f2b::/64 any eq ssh
 112 permit tcp 2001:4888:a03:1f2b::/64 any eq ssh
 113 permit tcp 2001:4888:a06:1f1b::/64 any eq ssh
 120 remark "iEN UCS"
 121 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq snmp
 122 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq snmp
 123 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq snmp
 124 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq 9980
 125 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq 9980
 126 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq 9980
 127 permit udp 2001:4888:a02:2100:a0:fef::/112 any eq 9990
 128 permit udp 2001:4888:a03:2201:c0:fef::/112 any eq 9990
 129 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq 9990
 130 remark "SyslogNG"
 131 permit udp host 2001:4888:a02:2109:a0:fef:0:84 any eq syslog
 132 permit udp host 2001:4888:a03:2109:c0:fef:0:84 any eq syslog
 140 remark "SevOne Servers and Pollers"
 141 permit udp 2001:4888:a06:1d52:f1:fef::/112 any eq snmp
 142 permit udp 2001:4888:a03:1d12:c0:fef::/112 any eq snmp
 143 permit udp 2001:4888:a02:1d12:a0:fef::/112 any eq snmp
 160 remark "Nokia NSP formerly SAM"
 161 permit udp host 2001:4888:a03:2114:c0:fef:0:2e any eq snmp
 162 permit udp host 2001:4888:a01:2114:a1:fef:0:2e any eq snmp
 170 remark "Syslog ULM"
 171 permit udp 2001:4888:a02:2101:a0:fef::/112 any eq syslog
 172 permit udp 2001:4888:a03:2217:c0:fef::/112 any eq syslog
 173 permit udp 2001:4888:a05:2203:e0:fef::/112 any eq syslog
 174 permit udp 2001:4888:a06:2281:f1:fef::/112 any eq syslog
 180 remark "SANE SSH"
 181 permit tcp 2001:4888:a01:2109:a1:9::/112 any eq ssh
 182 permit tcp 2001:4888:a01:2100:a1:fef::/112 any eq ssh
 183 permit tcp 2001:4888:a03:2144:c0:9::/112 any eq ssh
 184 permit tcp 2001:4888:a03:2100:c0:fef::/112 any eq ssh
 185 permit tcp 2001:4888:a06:2144:f0:9::/112 any eq ssh
 186 permit tcp 2001:4888:a06:2100:f0:fef::/112 any eq ssh
 200 remark "HPNA"
 201 permit tcp host 2001:4888:a02:2104:a0:fef:0:21 any eq ssh
 202 permit tcp host 2001:4888:a03:2110:c0:fef:0:12 any eq ssh
 210 remark "APSN NSOAUTO Service Portal"
 211 permit tcp host 2001:4888:a03:210c:c0:fef:0:20 any eq ssh
 220 remark "Smart Licensing On-Prem SSM"
 221 permit tcp host 2001:4888:a06:1f1b:f1:ff2:0:7 eq https any
 230 remark "EBH-AP"
 231 permit tcp host 2607:f160:8a05:1028:8000::6 any
 240 remark "Cisco CX Cloud"
 241 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc1 any
 242 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc2 any
 243 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc3 any
 244 permit tcp host 2001:4888:a26:2004:240:2c0:0:ccc4 any
 400 remark "Labs and FOA Sites Only"
 410 remark "EDN Desktops"
 411 permit tcp 2001:4888:a610::/44 any eq ssh
 412 permit tcp 2001:4888:a620::/44 any eq ssh
 413 permit tcp 2001:4888:a630::/44 any eq ssh
 414 permit tcp 2001:4888:a640::/44 any eq ssh
 415 permit tcp 2001:4888:a650::/44 any eq ssh
 416 permit tcp 2001:4888:a660::/44 any eq ssh
 417 permit tcp 2600:80b:150::/44 any eq ssh
 418 permit tcp 2600:80b:160::/44 any eq ssh
 419 permit tcp 2600:80b:170::/44 any eq ssh
 420 permit tcp 2600:80b:180::/44 any eq ssh
 421 permit tcp 2600:80b:190::/44 any eq ssh
 422 permit tcp 2600:80b:1a0::/44 any eq ssh
 423 permit tcp 2600:80b:1b0::/44 any eq ssh
 424 permit tcp 2600:80b:1c0::/44 any eq ssh
 450 remark "EDN VPN Pools"
 451 permit tcp 2001:4888:a600::/44 any eq ssh
 452 permit tcp 2600:80b:310::/44 any eq ssh
 453 permit tcp 2600:80b:300::/44 any eq ssh
 500 remark VZW Standard SNMP ACL
 501 permit ipv6 2001:4888:a01:2100::/56 any
 502 permit ipv6 2001:4888:a02:2100::/56 any
 503 permit ipv6 2001:4888:a03:2100::/56 any
 504 permit ipv6 2001:4888:a04:2100::/56 any
 505 permit ipv6 2001:4888:a05:2100::/56 any
 506 permit ipv6 2001:4888:a06:2100::/56 any
 507 permit ipv6 2001:4888:a07:2100::/56 any
 508 permit ipv6 2001:4888:a08:2100::/56 any
 509 permit ipv6 2001:4888:a0e:2100::/56 any
 510 permit ipv6 2001:4888:a0f:2100::/56 any
 511 permit ipv6 2001:4888:a06:2200::/56 any
 512 permit ipv6 2001:4888:a02:2200::/56 any
 513 permit ipv6 2001:4888:a02:1d10::/60 any
 514 permit ipv6 2001:4888:a06:1d50::/60 any
 515 permit ipv6 2001:4888:a03:1d10::/60 any
 516 permit ipv6 2001:4888:2:1d10::/60 any
 517 permit ipv6 2001:4888:6:1d50::/60 any
 518 permit ipv6 2001:4888:3:1d10::/60 any
 1000 deny ipv6 any any
!
ipv6 access-list SEV1_ACLv6
 10 remark Version_Q22023
 20 permit ipv6 2001:4888:a02:1d10::/60 any
 30 permit ipv6 2001:4888:a06:1d50::/60 any
 40 permit ipv6 2001:4888:a03:1d10::/60 any
 50 permit ipv6 2001:4888:2:1d10::/60 any
 60 permit ipv6 2001:4888:6:1d50::/60 any
 70 permit ipv6 2001:4888:3:1d10::/60 any
 80 deny ipv6 any any
!
ipv6 access-list SNMP_ACLv6
 10 remark Version_Q22023
 20 permit ipv6 2001:4888:a01:2100::/56 any
 30 permit ipv6 2001:4888:a02:2100::/56 any
 40 permit ipv6 2001:4888:a03:2100::/56 any
 50 permit ipv6 2001:4888:a03:2200::/56 any
 60 permit ipv6 2001:4888:a04:2100::/56 any
 70 permit ipv6 2001:4888:a05:2100::/56 any
 80 permit ipv6 2001:4888:a06:2100::/56 any
 90 permit ipv6 2001:4888:a07:2100::/56 any
 100 permit ipv6 2001:4888:a08:2100::/56 any
 110 permit ipv6 2001:4888:a0e:2100::/56 any
 120 permit ipv6 2001:4888:a0f:2100::/56 any
 130 permit ipv6 2001:4888:a06:2200::/56 any
 140 permit ipv6 2001:4888:a02:2200::/56 any
 150 permit ipv6 2001:4888:A03:2200::/56 any
 160 permit ipv6 2001:4888:a02:1d10::/60 any
 170 permit ipv6 2001:4888:a06:1d50::/60 any
 180 permit ipv6 2001:4888:a03:1d10::/60 any
 190 permit ipv6 2001:4888:2:1d10::/60 any
 200 permit ipv6 2001:4888:6:1d50::/60 any
 210 permit ipv6 2001:4888:3:1d10::/60 any
 220 deny ipv6 any any
!
ipv4 access-list MGMT_IN
 10 remark no IPv4 management access
 20 deny ipv4 any any
!
'''
            # Record all interfaces that backhaul to a HUB AL.
            uplink_hub_intf_list = [x for x in host_data["Interfaces"]
                                if host_data["Interfaces"][x]["description"].split("_")[0] in ciq_db
                                and host_data["eNSESR Role"] != ciq_db[host_data["Interfaces"][x]["description"].split("_")[0]]["eNSESR Role"]]
            # Record all interfaces that backhaul to a DRAN SPOKE.
            uplink_chain_intf_list = [x for x in host_data["Interfaces"]
                                if host_data["Interfaces"][x]["description"].split("_")[0] in ciq_db
                                and host_data["eNSESR Role"] == ciq_db[host_data["Interfaces"][x]["description"].split("_")[0]]["eNSESR Role"]]
            # If the DRAN SPOKE doesn't backhaul to a HUB AL, check that it backhauls to another DRAN SPOKE.
            if len(uplink_hub_intf_list) == 0:
                # Verify that the DRAN SPOKE has a backhaul interface.
                if len(uplink_chain_intf_list) == 0:
                    logger.critical(f"{host_data['New Hostname']} - Device has no uplink interface(s)")
                # If there is only 1 backhaul interface, there is no need for an ACL.
                elif len(uplink_chain_intf_list) == 1:
                    return conf
                elif len(uplink_chain_intf_list) > 1:
                    logger.critical(f"{host_data['New Hostname']} - Device is daisy-chain and has more than 1 uplink interface")
            # If there is only 1 backhaul interface, there is no need for an ACL.
            elif len(uplink_hub_intf_list) == 1:
                return conf
            # Device is dual-homed.
            elif len(uplink_hub_intf_list) == 2:
                return conf
            # Device has more than 2 connections to the HUB AL. This should not exist.
            elif len(uplink_hub_intf_list) > 2:
                logger.critical(f"{host_data['New Hostname']} - Device has more than 2 uplink interface(s)")
            return conf

        def qos():
            # MOP Section: QOS
            conf = f'''
class-map match-any WPS-IN
 match mpls experimental topmost 7
 match dscp 44
 end-class-map
!
class-map match-any GOLD-IN
 match mpls experimental topmost 3
 match dscp cs3 af21 af22 af23 af31 af32 af33
 end-class-map
!
class-map match-any VOICE-IN
 match mpls experimental topmost 5
 match dscp ef
 end-class-map
!
class-map match-any BRONZE-IN
 match mpls experimental topmost 1
 match dscp cs1
 end-class-map
!
class-map match-any SIGNAL-IN
 match mpls experimental topmost 4
 match dscp cs4 af41 af42 af43
 end-class-map
!
class-map match-any SILVER-IN
 match mpls experimental topmost 2
 match dscp cs2 af11 af12 af13
 end-class-map
!
class-map match-any WPS-QUEUE
 match traffic-class 7
 end-class-map
!
class-map match-any GOLD-QUEUE
 match traffic-class 3
 end-class-map
!
class-map match-any CTRL-BFD-IN
 match mpls experimental topmost 6
 match dscp cs6 cs7
 match precedence 6 7
 end-class-map
!
class-map match-any VOICE-QUEUE
 match traffic-class 5
 end-class-map
!
class-map match-any BRONZE-QUEUE
 match traffic-class 1
 end-class-map
!
class-map match-any SIGNAL-QUEUE
 match traffic-class 4
 end-class-map
!
class-map match-any SILVER-QUEUE
 match traffic-class 2
 end-class-map
!
class-map match-any CTRL-BFD-QUEUE
 match traffic-class 6
 end-class-map
!
policy-map QUEUES-OUT
 class BRONZE-QUEUE
  bandwidth remaining percent 10
  queue-limit 800 ms
 !
 class WPS-QUEUE
  shape average percent 5
  queue-limit 1 ms
  priority level 2
 !
 class CTRL-BFD-QUEUE
  shape average percent 2
  queue-limit 1 ms
  priority level 1
 !
 class VOICE-QUEUE
  shape average percent 60
  priority level 2
  queue-limit 1 ms
 !
 class SIGNAL-QUEUE
  bandwidth remaining percent 15
 !
 class GOLD-QUEUE
  bandwidth remaining percent 10
 !
 class SILVER-QUEUE
  bandwidth remaining percent 10
 !
 class class-default
  bandwidth remaining percent 37
  queue-limit 400 ms
 !
 end-policy-map
!
policy-map CLASSIFY-IN
 class BRONZE-IN
  set traffic-class 1
  set mpls experimental imposition 1
 !
 class WPS-IN
  set traffic-class 7
  set mpls experimental imposition 7
 !
 class CTRL-BFD-IN
  set traffic-class 6
  set mpls experimental imposition 6
 !
 class VOICE-IN
  set mpls experimental imposition 5
  set traffic-class 5
 !
 class SIGNAL-IN
  set mpls experimental imposition 4
  set traffic-class 4
 !
 class GOLD-IN
  set mpls experimental imposition 3
  set traffic-class 3
 !
 class SILVER-IN
  set mpls experimental imposition 2
  set traffic-class 2
 !
 class class-default
  set mpls experimental imposition 0
  set traffic-class 0
 !
 end-policy-map
!
policy-map CLASSIFY-CARRIER-AGG-IN
 class class-default
  set traffic-class 5
  set qos-group 5
 !
 end-policy-map
!
'''
            quads = [x.text for x in runfile.find_objects("hw-module quad")]
            if quads:
                for quad in quads:
                    quad_lines = [x for x in runfile.find_all_children(quad)]
                    for line in quad_lines:
                        conf += f'''{line}\n'''
                    conf += f'''!\n'''
            return conf

        def interfaces_loopback_and_management():
            # MOP Section: Interfaces - Loopback and Management
            conf = f'''
interface Loopback0
 description Global Loopback
 ipv4 address {host_data["Loopback 0 Global - IPv4"]}/32
!
interface Loopback1
 vrf RAN
 ipv6 address {host_data["Loopback 1 RAN - IPv6"]}/128
!
interface Loopback4
 vrf CELL_MGMT
 ipv6 address {host_data["Loopback 4 CELL_MGMT - IPv6"]}/128
!
interface MgmtEth0/RP0/CPU0/0
 vrf management
 ipv6 address {host_data["MGMT - IPv6"]}/64
 ipv4 access-group MGMT_IN ingress
 ipv6 access-group MGMT_IN_V6 ingress
!
'''
            return conf

        def interfaces_physical(prev_conf):
            conf = ""
            # MOP Section: Interfaces - Physical
            # Create a list of interfaces that will be configured based on the CIQ. This will be used when importing legacy interfaces.
            excluded_intf = []
            # Build new interfaces and shapers as needed.
            for intf in host_data["Interfaces"]:
                excluded_intf.append(intf.split('.')[0])
                # Capture the shorthand of the interface. Ex. Gi, Te, or Hu
                intf_sh = intf[:2].lower()
                circuit_rate = int(host_data["Interfaces"][intf]["circuit_rate"])
                if "B4A" in host_data["Interfaces"][intf]["description"].split('_')[0]:
                    if "Breakout" in host_data["Interfaces"][intf].keys():
                        if f'''controller Optics{host_data["Interfaces"][intf]["Breakout"]}''' not in conf and f'''controller Optics{host_data["Interfaces"][intf]["Breakout"]}''' not in prev_conf:
                            conf += f'''
controller optics {host_data["Interfaces"][intf]["Breakout"]}
 breakout 4x10
!
'''
                    if intf_sh == "gi":
                        if circuit_rate <= 0 or circuit_rate > 1000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 1000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 1000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 !
 end-policy-map
!
'''
                        elif circuit_rate == 1000:
                            intf_bw = 1000000
                            egress_policy = f'''QUEUES-OUT'''
                    elif intf_sh == "te":
                        if circuit_rate <= 0 or circuit_rate > 10000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 10000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 10000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 !
 end-policy-map
!
'''
                        elif circuit_rate == 10000:
                            intf_bw = 10000000
                            egress_policy = f'''QUEUES-OUT'''
                    elif intf_sh == "hu":
                        if circuit_rate <= 0 or circuit_rate > 100000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 100000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 100000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 !
 end-policy-map
!
'''
                        elif circuit_rate == 100000:
                            intf_bw = 100000000
                            egress_policy = f'''QUEUES-OUT'''
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                        exit()
                    if "clock-type T-BC" in prev_conf:
                        conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["cid"]}
 mtu 9100
 load-interval 30
!
interface {intf}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 bandwidth {intf_bw}
 load-interval 30
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 encapsulation dot1q {host_data["Interfaces"][intf]["VLAN"]}
 ipv4 point-to-point
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
 ptp
  profile slave
  transport ethernet
  port state slave-only
 frequency synchronization
  selection input
  wait-to-restore 0
!
'''
                    else:
                        conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["cid"]}
 mtu 9100
 load-interval 30
!
interface {intf}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 bandwidth {intf_bw}
 load-interval 30
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 encapsulation dot1q {host_data["Interfaces"][intf]["VLAN"]}
 ipv4 point-to-point
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
!
'''
                if "B4C" in host_data["Interfaces"][intf]["description"].split('_')[0]:
                    if "Breakout" in host_data["Interfaces"][intf].keys():
                        if f'''controller Optics{host_data["Interfaces"][intf]["Breakout"]}''' not in conf and f'''controller Optics{host_data["Interfaces"][intf]["Breakout"]}''' not in prev_conf:
                            conf += f'''
controller optics {host_data["Interfaces"][intf]["Breakout"]}
 breakout 4x10
!
'''
                    if intf_sh == "gi":
                        if circuit_rate <= 0 or circuit_rate > 1000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 1000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 1000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 !
 end-policy-map
!
'''
                        elif circuit_rate == 1000:
                            intf_bw = 1000000
                            egress_policy = f'''QUEUES-OUT'''
                    elif intf_sh == "te":
                        if circuit_rate <= 0 or circuit_rate > 10000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 10000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 10000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 !
 end-policy-map
!
'''
                        elif circuit_rate == 10000:
                            intf_bw = 10000000
                            egress_policy = f'''QUEUES-OUT'''
                    elif intf_sh == "hu":
                        if circuit_rate <= 0 or circuit_rate > 100000:
                            logger.critical(f"{host_data['New Hostname']} - Circuit Rate for interface {intf} is invalid. Range is 0 - 100000. Update CIQ and re-run script.")
                            exit()
                        elif 0 < circuit_rate < 100000:
                            shape_rate = math.floor(circuit_rate * 0.96)
                            intf_bw = shape_rate * 1000
                            egress_policy = f'''policy-map {circuit_rate}MB-NE'''
                            if f'''policy-map {circuit_rate}MB-NE''' not in conf:
                                conf += f'''
policy-map {circuit_rate}MB-NE
 class class-default
  service-policy QUEUES-OUT
  shape average {shape_rate} mbps
 !
 end-policy-map
!
'''
                        elif circuit_rate == 100000:
                            intf_bw = 100000000
                            egress_policy = f'''QUEUES-OUT'''
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Interface {intf} needs to start with Gi, Te, or Hu. Update CIQ and re-run script.")
                        exit()
                    if "clock-type T-BC" in prev_conf:
                        conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["cid"]}
 mtu 9100
 load-interval 30
!
interface {intf}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 bandwidth {intf_bw}
 load-interval 30
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 encapsulation dot1q {host_data["Interfaces"][intf]["VLAN"]}
 ipv4 point-to-point
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
 ptp
  profile slave
  transport ethernet
  port state slave-only
 frequency synchronization
  selection input
  wait-to-restore 0
!
'''
                    else:
                        conf += f'''
interface {intf.split(".")[0]}
 no shutdown
 description {host_data["Interfaces"][intf]["cid"]}
 mtu 9100
 load-interval 30
!
interface {intf}
 no shutdown
 description {host_data["Interfaces"][intf]["description"]}
 bandwidth {intf_bw}
 load-interval 30
 service-policy input CLASSIFY-IN
 service-policy output {egress_policy}
 encapsulation dot1q {host_data["Interfaces"][intf]["VLAN"]}
 ipv4 point-to-point
 ipv4 address {host_data["Interfaces"][intf]["ip"]}/31
!
'''

            # Exclude OOB interface
            excluded_intf += ["interface GigabitEthernet0/0/0/23", "interface GigabitEthernet0/0/0/23.400 l2transport"]
            conf += f'''
interface GigabitEthernet0/0/0/23
 no shutdown
 description CSR OOB Management
 negotiation auto
 load-interval 30
!
interface GigabitEthernet0/0/0/23.400 l2transport
 no shutdown
 encapsulation untagged
!
'''

            # Exclude any Bundles from being carried over.
            bundle_list = [x.text.split(".")[0] for x in runfile.find_objects(r"^interface.+Bundle.+") if "preconfigure" not in x.text]
            excluded_intf += bundle_list

            # Create a list of all main interfaces currently configured on the router.
            legacy_intf = [x.text for x in runfile.find_objects(r"^interface.+Gig.+") if "preconfigure" not in x.text]

            # Remove duplicate entries
            legacy_intf_sorted = sorted(set(legacy_intf), key=legacy_intf.index)
            legacy_intf_main_ports = [i for i in legacy_intf_sorted if "." not in i]

            # Create list of interfaces that are currently configured, but aren't listed to be configured in the CIQ.
            excluded_intf_main_ports = [re.findall(r"([\d/]+)", i)[0] for i in excluded_intf if i not in bundle_list]
            excluded_intf_sub_ports = [re.findall(r"([\d/.]+)", i)[0] for i in excluded_intf if i not in bundle_list]
            carried_intf = [l_intf for l_intf in legacy_intf_sorted
                            if re.findall(r"([\d/]+)", l_intf)[0] not in excluded_intf_main_ports
                            and re.findall(r"[\d/]+\.([\d]+).+", l_intf) not in excluded_intf_sub_ports
                            and all(re.findall(r"([\d/]+)", l_intf)[0] != re.findall(r"([\d/]+)", x_intf)[0] for x_intf in excluded_intf_main_ports)]
            carried_intf_copy = copy.deepcopy(carried_intf)
            for intf in carried_intf_copy:
                if "." in intf:
                    if re.findall(r"(.+)\.", intf)[0] in legacy_intf_main_ports:
                        continue
                    else:
                        carried_intf.remove(intf)

            # Carry over any currently used shapers.
            for intf in carried_intf:
                try:
                    intf_egress_policy = runfile.find_children_w_parents(intf, "service-policy output")
                except:
                    continue
                else:
                    policies = [x.split(" service-policy output ")[1] for x in intf_egress_policy if "MARK-OUT" not in x]
                    for policy in policies:
                        # If the policy was already added to the config, go to the next policy.
                        if f"policy-map {policy}" in conf or f"policy-map {policy}" in prev_conf:
                            continue
                        elif "-48MB" in policy:
                            if f"policy-map {policy.replace('-48MB', '-NE')}" in conf:
                                continue
                        elif "-46MB" in policy:
                            if f"policy-map {policy.replace('-46MB', '-ST')}" in conf:
                                continue
                        # Get all the config for the egress policy.
                        policy_lines = [x for x in runfile.find_all_children(f"^policy-map {policy}$")]
                        for line in policy_lines:
                            if "-48MB" in line:
                                conf += f'''{line.replace("-48MB", "-NE")}\n'''
                            elif "-46MB" in line:
                                conf += f'''{line.replace("-46MB", "-ST")}\n'''
                            else:
                                conf += f'''{line}\n'''
                        conf += f'''!\n'''

            shut_main_intf_list = [intf for intf in carried_intf if "." not in intf and "shutdown" in [x.strip() for x in runfile.find_all_children(f"^{intf}$")]]

            # Carry over the interface configs.
            for intf in carried_intf:
                shut = False
                if "." in intf:
                    if intf.split(".")[0] in shut_main_intf_list:
                        shut = True
                else:
                    if intf in shut_main_intf_list:
                        shut = True
                intf_lines = [x for x in runfile.find_all_children(f"^{intf}$")]
                for line in intf_lines:
                    # Don't keep the bundle command if it's present.
                    if "bundle" in line.lower():
                        continue
                    # Don't keep the MARK-OUT policy if it's present.
                    elif "mark-out" in line.lower():
                        continue
                    # Change -48MB to -NE
                    elif "-48mb" in line.lower():
                        conf += f'''{line.replace("-48MB", "-NE")}\n'''
                    # Change -46MB to -NE
                    elif "-46mb" in line.lower():
                        conf += f'''{line.replace("-46MB", "-ST")}\n'''
                    elif "interface" in line.lower():
                        conf += f"{line}\n"
                        if not shut:
                            conf += f" no shutdown\n"
                    else:
                        conf += f"{line}\n"
                conf += f"!\n"
            return conf

        def interfaces_bvi():
            conf = ""
            # MOP Section: Interfaces - BVI
            # Create a list of all possible legacy BVIs that were reserved for the HUB AL connections.
            excluded_bvi = ["interface BVI300", "interface BVI400"]
            if host_data["BVI100 - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI100 - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI100 - IPv6"]]:
                excluded_bvi.append(f"interface BVI300")
                if host_data["BVIs need suppress-ra?"] is not None:
                    if "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                        conf += f'''
interface BVI100
 description eNodeB
 host-routing
 vrf RAN
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                        for ip_addr in host_data["BVI100 - IPv6"]:
                            conf += f'''
 ipv6 address {ip_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                    elif "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                        conf += f'''
interface BVI100
 description eNodeB
 host-routing
 vrf RAN
 ipv6 mtu 1970
'''
                        for ip_addr in host_data["BVI100 - IPv6"]:
                            conf += f'''
 ipv6 address {ip_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Samsung VDU site must be Yes or No in CIQ.")
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            else:
                logger.critical(f"{host_data['New Hostname']} - BVI100 IP not provided in CIQ.")
            if host_data["BVI400 - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI400 - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI400 - IPv6"]]:
                excluded_bvi.append(f"interface BVI400")
                if host_data["BVIs need suppress-ra?"] is not None:
                    if "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                        conf += f'''
interface BVI400
 description CELL_MGMT VLAN interface
 host-routing
 vrf CELL_MGMT
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                        for ip_addr in host_data["BVI400 - IPv6"]:
                            conf += f'''
 ipv6 address {ip_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                    elif "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                        conf += f'''
interface BVI400
 description CELL_MGMT VLAN interface
 host-routing
 vrf CELL_MGMT
 ipv6 mtu 1970
'''
                        for ip_addr in host_data["BVI400 - IPv6"]:
                            conf += f'''
 ipv6 address {ip_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Column 'BVIs need suppress-ra?' needs to be filled out with yes/no.")
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            else:
                logger.critical(f"{host_data['New Hostname']} - BVI400 IP not provided in CIQ.")
            if host_data["BVI150 - IPv6"] is not None and "n/a" not in [x.lower() for x in host_data["BVI150 - IPv6"]] and "na" not in [x.lower() for x in host_data["BVI150 - IPv6"]]:
                excluded_bvi.append(f"interface BVI150")
                excluded_bvi.append(f"interface BVI310")
                if host_data["BVIs need suppress-ra?"] is not None:
                    if "yes" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                        conf += f'''
interface BVI150
 description eNodeB
 host-routing
 vrf RAN
 ipv6 nd suppress-ra
 ipv6 mtu 1970
'''
                        for ip_addr in host_data["BVI150 - IPv6"]:
                            conf += f'''
 ipv6 address {ip_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                    elif "no" in [x.lower() for x in host_data["BVIs need suppress-ra?"]]:
                        conf += f'''
interface BVI150
 description eNodeB
 host-routing
 vrf RAN
 ipv6 mtu 1970
'''
                        for ip_addr in host_data["BVI150 - IPv6"]:
                            conf += f'''
 ipv6 address {ip_addr}/64
'''
                        conf += f'''
 load-interval 30
!
'''
                    else:
                        logger.critical(f"{host_data['New Hostname']} - Samsung VDU site must be Yes or No in CIQ.")
                else:
                    logger.critical(f"{host_data['New Hostname']} - Samsung VDU site not specified in CIQ.")
            # Grab any existing BVIs from the DRAN SPOKE
            legacy_bvi = [x.text.split(".")[0] for x in runfile.find_objects(r"^interface.+BVI.+")]
            legacy_bvi.append("interface BVI300")
            legacy_bvi_sorted = sorted(set(legacy_bvi), key=legacy_bvi.index)
            carry_bvi = [x for x in legacy_bvi_sorted if x not in excluded_bvi]
            for bvi in carry_bvi:
                bvi_lines = runfile.find_all_children(f"^{bvi}")
                hr_check = [x for x in bvi_lines if "host-routing" in x]

                # Update the BVI VRF if needed.
                for line in bvi_lines:
                    if "LTE" in line:
                        bvi_lines[bvi_lines.index(line)] = line.replace("LTE", "RAN")

                # Add host routing if the BVI didn't have it before.
                if not hr_check:
                    bvi_lines.insert(1, " host-routing")

                # Add the BVI to the config
                for line in bvi_lines:
                    conf += f'''{line}\n'''
                conf += "!\n"
            return conf

        def l2vpn(new_conf):
            conf = ""
            # MOP Section: L2VPN
            # Get list of existing bridge domains.
            legacy_bd = [x for x in runfile.find_children_w_parents("^ bridge group", "^  bridge-domain")]
            # Instantiate lists that will hold configs
            legacy_bd_300 = []
            legacy_bd_310 = []
            legacy_bd_400 = []
            remaining_bd = []
            for bd in legacy_bd:
                # Grab any existing interfaces from the different bridge domains if they exist.
                if "bridge-domain 350" in bd.strip():
                    continue
                elif "bridge-domain 300" in bd.strip():
                    legacy_bd_300 = [x for x in runfile.find_children_w_parents("^  bridge-domain 300$", "^   interface")]
                elif "bridge-domain 310" in bd.strip():
                    legacy_bd_310 = [x for x in runfile.find_children_w_parents("^  bridge-domain 310$", "^   interface")]
                elif "bridge-domain 400" in bd.strip():
                    legacy_bd_400 = [x for x in runfile.find_children_w_parents("^  bridge-domain 400$", "^   interface")]
                else:
                    remaining_bd.append(bd)

            # Add L2VPN config for vrf RAN
            conf += f'''
l2vpn
 bridge group xNB-BG'''
            if "interface BVI100" in new_conf:
                conf += f'''
  bridge-domain 100'''
                # Add config for BD 100 using legacy BD 300 config if it existed.
                if len(legacy_bd_300) > 0:
                    for intf in legacy_bd_300:
                        intf_num = re.findall(r"([\d/.]+)", intf)[0]
                        if "bundle" in intf.lower():
                            continue
                        elif re.search(f'''\ninterface.+{intf_num}.*\n''', new_conf):
                            conf += f'''
   {intf.strip()}
   !'''
                conf += f'''
   routed interface BVI100
   !
  !'''
            else:
                logger.critical(f"{host_data['New Hostname']} - BVI100 wasn't in 'conf', so the L2VPN can't be built.")

            if "interface BVI150" in new_conf:
                conf += f'''
  bridge-domain 150'''
                # Add config for BD 150 using legacy BD 310 config if it existed.
                if len(legacy_bd_310) > 0:
                    for intf in legacy_bd_310:
                        intf_num = re.findall(r"([\d/.]+)", intf)[0]
                        if "bundle" in intf.lower():
                            continue
                        elif re.search(f'''\ninterface.+{intf_num}.*\n''', new_conf):
                            conf += f'''
   {intf.strip()}
   !'''
                conf += f'''
   routed interface BVI150
   !
  !'''

            # Add config for BD 400.
            if f"interface BVI400" in new_conf:
                conf += f'''
  bridge-domain 400'''
                if f"\ninterface GigabitEthernet0/0/0/23.400 l2transport\n" in new_conf:
                    conf += f'''
   interface GigabitEthernet0/0/0/23.400
   !'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - Port 0/0/0/23 is already configured. The L2VPN config will not be applied.")
                if len(legacy_bd_400) > 0:
                    for intf in legacy_bd_400:
                        intf_num = re.findall(r"([\d/.]+)", intf)[0]
                        if "bundle" in intf.lower():
                            continue
                        elif re.search(f'''\ninterface.+{intf_num}.*\n''', new_conf):
                            conf += f'''
   {intf.strip()}
   !'''
                conf += f'''
   routed interface BVI400
   !
  !
'''
            else:
                logger.critical(f"{host_data['New Hostname']} - BVI400 wasn't in 'conf', so the L2VPN can't be built.")

            # Add any remaining legacy BDs into CELL_MGMT
            if len(remaining_bd) > 0:
                for bd in remaining_bd:
                    legacy_bd_config = [x for x in runfile.find_all_children(f"^{bd}$")]
                    bd_bvi_num = [re.search(r"bvi (\d+)$", x) for x in legacy_bd_config if "bvi" in x]
                    if bd_bvi_num:
                        if f"interface BVI{bd_bvi_num}" in new_conf:
                            for line in legacy_bd_config:
                                if "bundle" in line.lower():
                                    continue
                                elif "interface" in line.lower():
                                    intf_num = re.findall(r"([\d/.]+)", intf)[0]
                                    if re.search(f'''\ninterface.+{intf_num}.*\n''', new_conf):
                                        conf += f'''{line}
    '''
                                else:
                                    conf += f'''{line}
    '''
                            conf += f'''  !
'''
                        else:
                            logger.critical(f"{host_data['New Hostname']} - BVI{bd_bvi_num} wasn't in 'conf', so the L2VPN can't be built.")
                    else:
                        for line in legacy_bd_config:
                            if "bundle" in line.lower():
                                continue
                            else:
                                conf += f'''{line}
'''
                        conf += f'''  !
'''

            conf += f'''
 !
!
'''
            return conf

        def prefix_sets():
            # MOP Section: Prefix-Sets
            conf = f'''
prefix-set PRFX_DEFAULT
  0.0.0.0/0,
  ::/0
end-set
!
'''
            return conf

        def route_policies():
            # MOP Section: Route-Policies
            conf = f'''
route-policy LABEL_LOOPBACK0
  set aigp-metric igp-cost
  set label-index {host_data["ISIS SR Prefix SID"]}
end-policy
!
route-policy SET_ADDPATH
  set path-selection backup 1 install multipath-protect advertise
end-policy
!'''
            if "l3vpn" not in [host_data["BGP_Peers"][x]["Solution"].lower() for x in host_data["BGP_Peers"]]:
                if "hub al" in [host_data["BGP_Peers"][x]["eNSESR Role"].lower() for x in host_data["BGP_Peers"]]:
                    conf += f'''
route-policy IMPORT_RR-5-ENSESR_CSR-AL_LABEL
 pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_CSR-AL_LABEL
 pass
end-policy
!
route-policy IMPORT_RR-5-ENSESR_CSR-AL_EVPN
 pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_CSR-AL_EVPN
 pass
end-policy
!
'''
                    if "dran spoke" in [host_data["BGP_Peers"][x]["eNSESR Role"].lower() for x in host_data["BGP_Peers"]]:
                        conf += f'''
route-policy IMPORT_RR-5-ENSESR_CSR-SPOKE_LABEL
 pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_CSR-SPOKE_LABEL
 pass
end-policy
!
route-policy IMPORT_RR-5-ENSESR_CSR-SPOKE_EVPN
 pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_CSR-SPOKE_EVPN
 pass
end-policy
!
'''
                else:
                    if "dran spoke" in [host_data["BGP_Peers"][x]["eNSESR Role"].lower() for x in host_data["BGP_Peers"]]:
                        conf += f'''
route-policy IMPORT_RR-5-ENSESR_SPOKE-CSR_LABEL
 pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_SPOKE-CSR_LABEL
 pass
end-policy
!
route-policy IMPORT_RR-5-ENSESR_SPOKE-CSR_EVPN
 pass
end-policy
!
route-policy EXPORT_RR-5-ENSESR_SPOKE-CSR_EVPN
 pass
end-policy
!
'''
            else:
                if "hub al" in [host_data["BGP_Peers"][x]["eNSESR Role"].lower() for x in host_data["BGP_Peers"]]:
                    conf += f'''
route-policy IMPORT_RR-5-L3VPN_CSR-AL_LABEL
 pass
end-policy
!
route-policy EXPORT_RR-5-L3VPN_CSR-AL_LABEL
 pass
end-policy
!
route-policy IMPORT_RR-5-L3VPN_CSR-AL_L3VPN
 pass
end-policy
!
route-policy EXPORT_RR-5-L3VPN_CSR-AL_L3VPN
 pass
end-policy
!
'''
                    if "dran spoke" in [host_data["BGP_Peers"][x]["eNSESR Role"].lower() for x in host_data["BGP_Peers"]]:
                        conf += f'''
route-policy IMPORT_RR-5-L3VPN_CSR-SPOKE_LABEL
 pass
end-policy
!
route-policy EXPORT_RR-5-L3VPN_CSR-SPOKE_LABEL
 pass
end-policy
!
route-policy IMPORT_RR-5-L3VPN_CSR-SPOKE_L3VPN
 pass
end-policy
!
route-policy EXPORT_RR-5-L3VPN_CSR-SPOKE_L3VPN
 pass
end-policy
!
'''
                else:
                    if "dran spoke" in [host_data["BGP_Peers"][x]["eNSESR Role"].lower() for x in host_data["BGP_Peers"]]:
                        conf += f'''
route-policy IMPORT_RR-5-L3VPN_SPOKE-CSR_LABEL
 pass
end-policy
!
route-policy EXPORT_RR-5-L3VPN_SPOKE-CSR_LABEL
 pass
end-policy
!
route-policy IMPORT_RR-5-L3VPN_SPOKE-CSR_L3VPN
 pass
end-policy
!
route-policy EXPORT_RR-5-L3VPN_SPOKE-CSR_L3VPN
 pass
end-policy
!
'''
            return conf

        def static_routes():
            # MOP Section: Static Routes
            conf = f'''
router static
 vrf management
  address-family ipv6 unicast
   ::/0 MgmtEth0/RP0/CPU0/0 {host_data["BVI400 - IPv6"][0]}
  !
 !
!
'''
            return conf

        def isis():
            conf = ""
            # MOP Section: ISIS Process 5
            # Add base ISIS config.
            conf += f'''
router isis 5
 set-overload-bit on-startup 180
 is-type level-1
 net {host_data["ISIS NET ID"]}
 nsf ietf
 log adjacency changes
 lsp-gen-interval maximum-wait 5000 initial-wait 50 secondary-wait 200
 lsp-refresh-interval 65000
 max-lsp-lifetime 65535
 lsp-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
 address-family ipv4 unicast
  metric-style wide
  microloop avoidance segment-routing
  mpls traffic-eng router-id Loopback0
  spf-interval maximum-wait 5000 initial-wait 50 secondary-wait 200
  segment-routing mpls sr-prefer
  spf prefix-priority critical tag 100
 !
 interface Loopback0
  address-family ipv4 unicast
   prefix-sid index {host_data["ISIS SR Prefix SID"]}
  !
 !'''

            # Add ISIS Interface config.
            for intf in host_data["Interfaces"]:
                if host_data["Interfaces"][intf]["description"].split("_")[0] in ciq_db.keys():
                    intf_metric = ""
                    if "HUB AL" in ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]["eNSESR Role"]:
                        intf_metric = "1000000"
                    elif "DRAN SPOKE" in ciq_db[host_data["Interfaces"][intf]["description"].split("_")[0]]["eNSESR Role"]:
                        intf_metric = "10"
                    conf += f'''
 interface {intf}
  circuit-type level-1
  bfd minimum-interval 50
  bfd multiplier 5
  bfd fast-detect ipv4
  point-to-point
  hello-padding disable
  hello-password hmac-md5 clear {host_data["ISIS PASSWORD"]}
  address-family ipv4 unicast
   fast-reroute per-prefix
   fast-reroute per-prefix ti-lfa
   metric {intf_metric}
  !
 !
'''
                else:
                    continue
            return conf

        def admin_part_3():
            # MOP Section: Admin Part 3
            conf = f'''
!
snmp-server traps isis all
snmp-server traps bgp cbgp2 updown
snmp-server traps bgp updown
'''
            return conf

        def bgp():
            conf = ""
            # MOP Section: BGP
            # Add base BGP config.
            conf += f'''
router bgp {host_data["COIN ASN"]}
 bgp router-id {host_data["Loopback 0 Global - IPv4"]}
 bgp graceful-restart
 ibgp policy out enforce-modifications
 !
 address-family ipv4 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
  network {host_data["Loopback 0 Global - IPv4"]}/32 route-policy LABEL_LOOPBACK0
  allocate-label all
 !
 address-family vpnv4 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
 address-family vpnv6 unicast
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
 address-family l2vpn evpn
  additional-paths receive
  additional-paths send
  additional-paths selection route-policy SET_ADDPATH
  nexthop trigger-delay critical 3000
  nexthop trigger-delay non-critical 10000
 !
'''
            # Determine which BGP NGs are required and build them.
            b4a_nbr = False
            b4c_nbr = False
            peer_ct = 0
            b4a_nbr_ct = 0
            b4c_nbr_ct = 0
            solutions = [host_data["BGP_Peers"][x]["Solution"] for x in host_data["BGP_Peers"]]
            for peer in host_data["BGP_Peers"]:
                peer_ct += 1
                if "B4A" in peer:
                    b4a_nbr = True
                    b4a_nbr_ct += 1
                elif "B4C" in peer:
                    b4c_nbr = True
                    b4c_nbr_ct += 1
            if peer_ct == 1:
                if b4a_nbr_ct == peer_ct:
                    conf += f'''
 neighbor-group RR-5-ENSESR_AL
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-policy IMPORT_RR-5-ENSESR_CSR-AL_LABEL in
   route-policy EXPORT_RR-5-ENSESR_CSR-AL_LABEL out
   next-hop-self
  !
  address-family l2vpn evpn
   route-policy IMPORT_RR-5-ENSESR_CSR-AL_EVPN in
   route-policy EXPORT_RR-5-ENSESR_CSR-AL_EVPN out
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
  !
 !
'''
                else:
                    if "L3VPN" in solutions:
                        conf += f'''
 neighbor-group RR-5-L3VPN_CSR
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   next-hop-self
   route-policy IMPORT_RR-5-L3VPN_SPOKE-CSR_LABEL in
   route-policy EXPORT_RR-5-L3VPN_SPOKE-CSR_LABEL out
  !
  address-family vpnv4 unicast
   route-policy IMPORT_RR-5-L3VPN_SPOKE-CSR_L3VPN in
   route-policy EXPORT_RR-5-L3VPN_SPOKE-CSR_L3VPN out
   advertise vpnv4 unicast re-originated
  !
  address-family vpnv6 unicast
   route-policy IMPORT_RR-5-L3VPN_SPOKE-CSR_L3VPN in
   route-policy EXPORT_RR-5-L3VPN_SPOKE-CSR_L3VPN out
   advertise vpnv6 unicast re-originated
  !
 !
'''
                    else:
                        conf += f'''
 neighbor-group RR-5-ENSESR_CSR
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-policy IMPORT_RR-5-ENSESR_SPOKE-CSR_LABEL in
   route-policy EXPORT_RR-5-ENSESR_SPOKE-CSR_LABEL out
   next-hop-self
  !
  address-family l2vpn evpn
   route-policy IMPORT_RR-5-ENSESR_SPOKE-CSR_EVPN in
   route-policy EXPORT_RR-5-ENSESR_SPOKE-CSR_EVPN out
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
  !
 !
'''
            elif peer_ct > 1:
                if b4a_nbr_ct == peer_ct:
                    conf += f'''
 neighbor-group RR-5-ENSESR_AL
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-policy IMPORT_RR-5-ENSESR_CSR-AL_LABEL in
   route-policy EXPORT_RR-5-ENSESR_CSR-AL_LABEL out
   next-hop-self
  !
  address-family l2vpn evpn
   route-policy IMPORT_RR-5-ENSESR_CSR-AL_EVPN in
   route-policy EXPORT_RR-5-ENSESR_CSR-AL_EVPN out
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
  !
 !
'''
                else:
                    if "L3VPN" in solutions:
                        conf += f'''
 neighbor-group RR-5-L3VPN_AL
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   next-hop-self
   route-policy IMPORT_RR-5-L3VPN_CSR-AL_LABEL in
   route-policy EXPORT_RR-5-L3VPN_CSR-AL_LABEL out
  !
  address-family vpnv4 unicast
   route-policy IMPORT_RR-5-L3VPN_CSR-AL_L3VPN in
   route-policy EXPORT_RR-5-L3VPN_CSR-AL_L3VPN out
   advertise vpnv4 unicast re-originated
  !
  address-family vpnv6 unicast
   route-policy IMPORT_RR-5-L3VPN_CSR-AL_L3VPN in
   route-policy EXPORT_RR-5-L3VPN_CSR-AL_L3VPN out
   advertise vpnv6 unicast re-originated
  !
 !
'''
                        conf += f'''
 neighbor-group RR-5-L3VPN_SPOKE
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-reflector-client
   next-hop-self
   route-policy IMPORT_RR-5-L3VPN_CSR-SPOKE_LABEL in
   route-policy EXPORT_RR-5-L3VPN_CSR-SPOKE_LABEL out
  !
  address-family vpnv4 unicast
   route-policy IMPORT_RR-5-L3VPN_CSR-SPOKE_L3VPN in
   route-policy EXPORT_RR-5-L3VPN_CSR-SPOKE_L3VPN out
   route-reflector-client
   advertise vpnv4 unicast re-originated
  !
  address-family vpnv6 unicast
   route-policy IMPORT_RR-5-L3VPN_CSR-SPOKE_L3VPN in
   route-policy EXPORT_RR-5-L3VPN_CSR-SPOKE_L3VPN out
   route-reflector-client
   advertise vpnv6 unicast re-originated
  !
 !
'''
                    else:
                        conf += f'''
 neighbor-group RR-5-ENSESR_AL
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-policy IMPORT_RR-5-ENSESR_CSR-AL_LABEL in
   route-policy EXPORT_RR-5-ENSESR_CSR-AL_LABEL out
   next-hop-self
  !
  address-family l2vpn evpn
   route-policy IMPORT_RR-5-ENSESR_CSR-AL_EVPN in
   route-policy EXPORT_RR-5-ENSESR_CSR-AL_EVPN out
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
  !
 !
'''
                        conf += f'''
 neighbor-group RR-5-ENSESR_SPOKE
  remote-as {host_data["COIN ASN"]}
  bfd fast-detect
  bfd multiplier 3
  bfd minimum-interval 100
  update-source Loopback0
  address-family ipv4 labeled-unicast
   route-policy IMPORT_RR-5-ENSESR_CSR-SPOKE_LABEL in
   route-reflector client
   route-policy EXPORT_RR-5-ENSESR_CSR-SPOKE_LABEL out
   next-hop-self
  !
  address-family l2vpn evpn
   route-policy IMPORT_RR-5-ENSESR_CSR-SPOKE_EVPN in
   route-reflector client
   route-policy EXPORT_RR-5-ENSESR_CSR-SPOKE_EVPN out
   advertise vpnv4 unicast re-originated
   advertise vpnv6 unicast re-originated
  !
 !
'''
            # Add BGP Neighbor config.
            for peer in host_data["BGP_Peers"]:
                if "HUB AL" in ciq_db[peer]["eNSESR Role"]:
                    if host_data["BGP_Peers"][peer]["Solution"] == "EVPN":
                        conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_AL
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                    elif host_data["BGP_Peers"][peer]["Solution"] == "L3VPN":
                        conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-L3VPN_AL
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                elif "DRAN SPOKE" in ciq_db[peer]["eNSESR Role"]:
                    if len(ciq_db[peer]["BGP_Peers"]) > 1:
                        if host_data["BGP_Peers"][peer]["Solution"] == "EVPN":
                            conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_CSR
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                        elif host_data["BGP_Peers"][peer]["Solution"] == "L3VPN":
                            conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-L3VPN_CSR
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                    else:
                        if host_data["BGP_Peers"][peer]["Solution"] == "EVPN":
                            conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-ENSESR_SPOKE
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                        elif host_data["BGP_Peers"][peer]["Solution"] == "L3VPN":
                            conf += f'''
 neighbor {host_data["BGP_Peers"][peer]["Loopback 0 Global - IPv4"]}
  use neighbor-group RR-5-L3VPN_SPOKE
  password clear {host_data["BGP PASSWORD"]}
  description {peer}
 !'''
                else:
                    logger.critical(f"{host_data['New Hostname']} - BGP Peer is not HUB AL or DRAN SPOKE.")
            # Add VRF config.
            conf += f'''
 vrf RAN
  rd {host_data["Loopback 0 Global - IPv4"]}:1
  bgp router-id {host_data["Loopback 0 Global - IPv4"]}
  address-family ipv6 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
 !
 vrf CELL_MGMT
  rd {host_data["Loopback 0 Global - IPv4"]}:4
  bgp router-id {host_data["Loopback 0 Global - IPv4"]}
  address-family ipv4 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
  address-family ipv6 unicast
   label mode per-vrf
   maximum-paths ibgp 16
   redistribute connected
  !
 !
!
'''
            return conf

        def admin_part_4():
            # MOP Section: Admin Part 4
            conf = f'''
snmp-server traps hsrp
snmp-server traps vrrp events
snmp-server traps l2vpn all
snmp-server traps l2vpn cisco
snmp-server traps l2vpn vc-up
snmp-server traps l2vpn vc-down
mpls oam
!
snmp-server traps mpls traffic-eng up
snmp-server traps mpls traffic-eng down
snmp-server traps mpls traffic-eng cisco
snmp-server traps mpls traffic-eng reroute
snmp-server traps mpls traffic-eng cisco-ext preempt
snmp-server traps mpls traffic-eng cisco-ext insuff-bw
snmp-server traps mpls traffic-eng cisco-ext bringup-fail
snmp-server traps mpls traffic-eng cisco-ext reroute-pending
snmp-server traps mpls traffic-eng cisco-ext reroute-pending-clear
snmp-server traps mpls traffic-eng reoptimize
snmp-server traps mpls frr all
snmp-server traps mpls frr protected
snmp-server traps mpls frr unprotected
snmp-server traps mpls ldp up
snmp-server traps mpls ldp down
snmp-server traps mpls ldp threshold
snmp-server traps mpls traffic-eng p2mp up
snmp-server traps mpls traffic-eng p2mp down
snmp-server traps mpls l3vpn all
snmp-server traps mpls l3vpn vrf-up
snmp-server traps mpls l3vpn vrf-down
snmp-server traps mpls l3vpn max-threshold-cleared
snmp-server traps mpls l3vpn max-threshold-exceeded
snmp-server traps mpls l3vpn mid-threshold-exceeded
segment-routing
 global-block 19000 119000
!
snmp-server traps sensor
snmp-server traps fru-ctrl
lldp
 management enable
 extended-show-width enable
!
snmp-server traps l2tun sessions
snmp-server traps l2tun tunnel-up
snmp-server traps l2tun tunnel-down
snmp-server traps l2tun pseudowire status
ssh client source-interface MgmtEth0/RP0/CPU0/0
ssh server enable cipher aes-cbc 3des-cbc
ssh timeout 30
ssh server rate-limit 100
ssh server session-limit 10
ssh server v2
ssh server vrf default
ssh server vrf management
ssh server vrf CELL_MGMT
snmp-server traps fabric plane
snmp-server traps fabric bundle link
snmp-server traps fabric bundle state
end
'''
            return conf

        logger.debug(f"Class: {self.__class__.__name__}")
        conf = ""
        conf += admin_part_1()
        conf += ptp()
        conf += vrfs()
        conf_temp = ""
        ntp_srvr1 = ""
        ntp_srvr2 = ""
        conf_temp, ntp_srvr1, ntp_srvr2 = admin_part_2(ntp_srvr1, ntp_srvr2)
        conf += conf_temp
        conf += acls(ntp_srvr1, ntp_srvr2)
        conf += qos()
        conf += interfaces_loopback_and_management()
        conf += interfaces_physical(conf)
        conf += interfaces_bvi()
        conf += l2vpn(conf)
        conf += prefix_sets()
        conf += route_policies()
        conf += static_routes()
        conf += isis()
        conf += admin_part_3()
        conf += bgp()
        conf += admin_part_4()

        # Remove unnecessary blank lines from the new config file.
        # Break up the config into 3 parts. Before banner motd, banner motd, and after banner motd
        conf_parts = conf.split("^", 2)
        conf_p1 = "\n".join([s for s in conf_parts[0].splitlines() if s])
        conf_p2 = conf_parts[1]
        conf_p3 = "\n".join([s for s in conf_parts[2].splitlines() if s])
        conf = "\n".join(["^".join([conf_p1, conf_p2, ""]), conf_p3])
        return conf


def read_ciq(ciq_file):
    # ciq_db is initialized and will be used to store device info from the CIQ.
    ciq_db = {}

    # Load the CIQ into a variable 'wb'.
    wb = load_workbook(ciq_file)

    # Specify the sheets that should be parsed during info gathering.
    req_sheets = [
        'Hub Info',
        'Site Configuration Data',
        'EBH AL to HUB BL P2P',
        'HUB BL to HUB BL P2P',
        'HUB BL to HUB AL P2P',
        'HUB AL to SPOKE P2P',
        'SPOKE to SPOKE P2P',
        'Non-RAN Services']
    # Add INFO log entry when the worksheet validation starts.
    logger.info("Workbook validation - Started")
    # Variable used to check if all sheets are in the CIQ.
    sheet_exists = True
    for sheet in req_sheets:
        # If the sheet doesn't exist in the CIQ, create a log entry and set the 'sheet_exists' variable.
        if sheet.lower() not in [x.lower() for x in wb.sheetnames]:
            logger.warning(f'''<{sheet}> is misspelled or missing''')
            sheet_exists = False
    # Create log entries depending on if all sheets were found, or if any were missing.
    if sheet_exists:
        logger.info("Workbook validation - Completed")
    else:
        logger.info("Workbook validation - Completed - Sheets are missing or misspelled")

    # Required headers are listed here based on the sheet they belong to.
    req_headers = {
        'Hub Info': [
            'Site Location',
            'Site Name',
            'COIN ASN',
            'ISIS 5 Password',
            'BGP Password',
            'Site Loopback0 Prefix',
            'EBH /31 Prefix for ODD GRE VLAN',
            'EBH /31 Prefix for EVEN GRE VLAN'],
        'Site Configuration Data': [
            'Migration Order',
            'Granite Site Name',
            'Old Hostname',
            'Old MGMT IP',
            'eNSESR Role',
            'Model',
            'Software Version',
            'New Hostname',
            'MGMT - IPv6',
            'Loopback 0 Global - IPv4',
            'Loopback 1 RAN - IPv6',
            'Loopback 4 CELL_MGMT - IPv6',
            'ISIS SR Prefix SID',
            'ISIS NET ID',
            'BVI100 - IPv6',
            'BVI150 - IPv6',
            'BVI400 - IPv6',
            'BVI450 - IPv4 CIDR',
            'BVI450 - IPv6',
            'BVI10X - IPv6',
            'BVI15X - IPv6',
            'BVI40X - IPv6',
            'BVI350 - IPv6',
            'BVIs need suppress-ra?'],
        'EBH AL to HUB BL P2P': [
            'EBH AL Hostname',
            'VLAN',
            'EBH AL Interface',
            'IPv4 /31 A',
            'IPv4 /31 Z',
            'HUB BL Hostname',
            'HUB BL Interface',
            'Circuit Rate (Mbps)',
            'Provider',
            'Circuit ID',
            'HUB BL Legacy Uplink - Far End Hostname',
            'HUB BL Legacy Uplink - Far End Interface',
            'HUB BL Legacy Uplink - Local Interface'],
        'HUB BL to HUB BL P2P': [
            'HUB BL Hostname A',
            'HUB BL Interface A',
            'IPv4 /31 A',
            'IPv4 /31 Z',
            'HUB BL Interface Z',
            'HUB BL Hostname Z'],
        'HUB BL to HUB AL P2P': [
            'HUB BL Hostname',
            'HUB BL Interface',
            'IPv4 /31 A',
            'IPv4 /31 Z',
            'HUB AL Interface',
            'HUB AL Hostname',
            'HUB AL Legacy Uplink - Far End Hostname',
            'HUB AL Legacy Uplink - Far End Interface',
            'HUB AL Legacy Uplink - Local Interface'],
        'HUB AL to SPOKE P2P': [
            'HUB AL Hostname',
            'VLAN',
            'HUB AL Interface',
            'IPv4 /31 A',
            'IPv4 /31 Z',
            'SPOKE Interface',
            'SPOKE Hostname',
            'Circuit Rate (Mbps)',
            'Provider',
            'Circuit ID',
            'SPOKE Legacy Uplink - Far End Hostname',
            'SPOKE Legacy Uplink - Far End Interface',
            'SPOKE Legacy Uplink - Local Interface'],
        'SPOKE to SPOKE P2P': [
            'Upstream SPOKE Hostname Connected to HUB AL',
            'Upstream SPOKE Hostname',
            'VLAN',
            'Upstream SPOKE Interface',
            'IPv4 /31 A',
            'IPv4 /31 Z',
            'Downstream SPOKE Interface',
            'Downstream SPOKE Hostname',
            'Circuit Rate (Mbps)',
            'Provider',
            'Circuit ID',
            'Downstream SPOKE Legacy Uplink - Far End Hostname',
            'Downstream SPOKE Legacy Uplink - Far End Interface',
            'Downstream SPOKE Legacy Uplink - Local Interface'],
        'Non-RAN Services': [
            'HUB Device Hostname',
            'HUB Device Interface',
            'Type',
            'Non-RAN Device Hostname',
            'Non-RAN Device Interface',
            'Non-RAN Device Legacy Uplink - Far End Hostname',
            'Non-RAN Device Legacy Uplink - Far End Interface',
            'Non-RAN Device Legacy Uplink - Local Interface']
    }
    # Go through each sheet and verify that the required headers are there.
    for sheet in req_headers:
        # Put the sheet data into variable 'ws'
        ws = wb[sheet]
        # Create a log indicating the start of the sheet's header validation check.
        logger.info(f'{ws} validation - Started')
        # Variable used to check if all headers are in the sheet.
        header_exists = True
        # Get the required headers for the sheet.
        headers = req_headers[sheet]
        # Record the headers actually configured in the sheet from the CIQ.
        h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]
        # Check that the required headers are configured in the sheet from the CIQ.
        for header in h_list:
            # If the header doesn't exist in the sheet, create a log entry and set the 'header_exists' variable.
            if header is not None and header.lower() not in [x.lower() for x in headers]:
                logger.info(f'''<{header}> is extra in {ws}''')
                header_exists = False
        # Create log entries depending on if all headers were found, or if any were missing.
        if header_exists:
            logger.info(f'{ws} validation - Completed')
        else:
            logger.info(f'{ws} validation - Completed - There are extra headers')

    # Record the site info from the Hub Info sheet.
    # hub_info_sheet = wb['Hub Info']
    ws = wb['Hub Info']
    headers = req_headers['Hub Info']
    h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]
    site_loc = ws.cell(2, h_list.index('Site Location') + 1).value      # Record the Site Location from cell A2 in sheet Hub Info
    site_name = ws.cell(2, h_list.index('Site Name') + 1).value         # Record the Site Name from cell B2 in sheet Hub Info
    coin_asn = ws.cell(2, h_list.index('COIN ASN') + 1).value           # Record the COIN ASN from cell C2 in sheet Hub Info
    isis_pw = ws.cell(2, h_list.index('ISIS 5 Password') + 1).value     # Record the ISIS Password from cell D2 in sheet Hub Info
    bgp_pw = ws.cell(2, h_list.index('BGP Password') + 1).value         # Record the BGP Password from cell E2 in sheet Hub Info
    site_loop0_prfx = ws.cell(2, h_list.index('Site Loopback0 Prefix') + 1).value         # Record the Site Loopback0 Prefix from cell F2 in sheet Hub Info
    ebh_gre_odd = ws.cell(2, h_list.index('EBH /31 Prefix for ODD GRE VLAN') + 1).value       # Record the GRE Prefix for ODD VLAN from cell G2 in sheet Hub Info
    ebh_gre_even = ws.cell(2, h_list.index('EBH /31 Prefix for EVEN GRE VLAN') + 1).value       # Record the GRE Prefix for EVEN VLAN from cell H2 in sheet Hub Info

    # Record data provided in the CIQ and add it to the ciq_db.
    for sheet in req_sheets:
        # Record data provided in the sheet and add it to ciq_db.
        match sheet:
            case 'Site Configuration Data':
                ws = wb[sheet]
                logger.info(f"Reading {ws} - Started")

                # Record the headers actually configured in the sheet from the CIQ.
                headers = req_headers[sheet]
                h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]

                # Go through each row and record data.
                for row in ws.iter_rows():
                    # If the 'Granite Site Name' is blank, stop processing data.
                    if row[h_list.index('Granite Site Name')].value is None:
                        break
                    # If the row has data and the row is not the headers, record the data.
                    if row[h_list.index('Granite Site Name')].row > 1:
                        # Build the device entry in ciq_db. Add all values from the required headers.
                        ciq_db[row[h_list.index('New Hostname')].value] = {x: row[h_list.index(x)].value for x in req_headers[sheet] if x in req_headers[sheet]}
                        copied_ciq_db = copy.deepcopy(ciq_db)
                        for entry in copied_ciq_db[row[h_list.index('New Hostname')].value]:
                            if "BVI" in entry:
                                data = copied_ciq_db[row[h_list.index('New Hostname')].value][entry]
                                if "\n" in data:
                                    conf_data = [x for x in data.split("\n")]
                                    del ciq_db[row[h_list.index('New Hostname')].value][entry]
                                    ciq_db[row[h_list.index('New Hostname')].value][entry] = conf_data
                                elif data is not None and "n/a" not in [x.lower() for x in data] and "na" not in [x.lower() for x in data]:
                                    ciq_db[row[h_list.index('New Hostname')].value][entry] = [data]
                        # Build the device entry in ciq_db. Add 'COIN ASN'.
                        ciq_db[row[h_list.index('New Hostname')].value]['SITE LOCATION'] = site_loc
                        ciq_db[row[h_list.index('New Hostname')].value]['SITE NAME'] = site_name
                        ciq_db[row[h_list.index('New Hostname')].value]['COIN ASN'] = coin_asn
                        ciq_db[row[h_list.index('New Hostname')].value]['ISIS PASSWORD'] = isis_pw
                        ciq_db[row[h_list.index('New Hostname')].value]['BGP PASSWORD'] = bgp_pw
                        ciq_db[row[h_list.index('New Hostname')].value]['SITE LOOPBACK0 PREFIX'] = site_loop0_prfx
                        ciq_db[row[h_list.index('New Hostname')].value]['EBH ODD GRE'] = ebh_gre_odd
                        ciq_db[row[h_list.index('New Hostname')].value]['EBH EVEN GRE'] = ebh_gre_even
                logger.info(f"Reading {ws} - Completed")

            # Record data provided in the sheet and add it to ciq_db.
            case 'EBH AL to HUB BL P2P':
                ws = wb[sheet]
                logger.info(f"Reading {ws} - Started")

                # Record the headers actually configured in the sheet from the CIQ.
                headers = req_headers[sheet]
                h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]

                # Go through each row and record data.
                for row in ws.iter_rows():
                    # If the 'EBH AL Hostname' is blank, stop processing data.
                    if row[h_list.index('EBH AL Hostname')].value is None:
                        break
                    # If the row has data and the row is not the headers, record the data.
                    if row[h_list.index('EBH AL Hostname')].row > 1:
                        # node1 is the EBH AL; node2 is the HUB BL
                        node1 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('EBH AL Hostname')].value.strip()][0]
                        node2 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('HUB BL Hostname')].value.strip()][0]
                        # Create an interface dict within the device entry in ciq_db
                        if 'Interfaces' not in ciq_db[node1]:
                            ciq_db[node1]['Interfaces'] = {}
                        if 'Interfaces' not in ciq_db[node2]:
                            ciq_db[node2]['Interfaces'] = {}
                        ebh_al_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('EBH AL Interface')].value)).group(1)
                        ebh_al_intf_word_full = ""
                        if "gi" in ebh_al_intf_word.lower()[:2]:
                            ebh_al_intf_word_full = "GigabitEthernet"
                        elif "te" in ebh_al_intf_word.lower()[:2]:
                            ebh_al_intf_word_full = "TenGigE"
                        elif "tw" in ebh_al_intf_word.lower()[:2]:
                            ebh_al_intf_word_full = "TwentyfiveGig"
                        elif "fo" in ebh_al_intf_word.lower()[:2]:
                            ebh_al_intf_word_full = "FortyGigE"
                        elif "hu" in ebh_al_intf_word.lower()[:2]:
                            ebh_al_intf_word_full = "HundredGigE"
                        hub_bl_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB BL Interface')].value)).group(1)
                        hub_bl_intf_word_full = ""
                        if "gi" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "GigabitEthernet"
                        elif "te" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "TenGigE"
                        elif "tw" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "TwentyfiveGig"
                        elif "fo" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "FortyGigE"
                        elif "hu" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "HundredGigE"
                        ebh_al_subintf = f'''{str(row[h_list.index('EBH AL Interface')].value).replace(ebh_al_intf_word, ebh_al_intf_word_full, 1)}.{row[h_list.index('VLAN')].value}'''
                        hub_bl_intf = f'''{str(row[h_list.index('HUB BL Interface')].value).replace(hub_bl_intf_word, hub_bl_intf_word_full, 1)}.{row[h_list.index('VLAN')].value}'''
                        ciq_db[node1]['Interfaces'][ebh_al_subintf] = {}
                        ciq_db[node2]['Interfaces'][hub_bl_intf] = {}
                        # Add the VLAN
                        ciq_db[node1]['Interfaces'][ebh_al_subintf]['VLAN'] = f'''{row[h_list.index('VLAN')].value}'''
                        ciq_db[node2]['Interfaces'][hub_bl_intf]['VLAN'] = f'''{row[h_list.index('VLAN')].value}'''
                        # Add the IPs under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][ebh_al_subintf]['ip'] = f'''{row[h_list.index('IPv4 /31 A')].value}'''
                        ciq_db[node2]['Interfaces'][hub_bl_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 Z')].value}'''
                        # Add the interface descriptions under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][ebh_al_subintf]['description'] = f'''{row[h_list.index('HUB BL Hostname')].value}_{row[h_list.index('HUB BL Interface')].value}'''
                        ciq_db[node2]['Interfaces'][hub_bl_intf]['description'] = f'''{row[h_list.index('EBH AL Hostname')].value}_{row[h_list.index('EBH AL Interface')].value}'''
                        # Add the circuit rate under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][ebh_al_subintf]['circuit_rate'] = ciq_db[node2]['Interfaces'][hub_bl_intf]['circuit_rate'] = f'''{row[h_list.index('Circuit Rate (Mbps)')].value}'''
                        # Add the circuit ID under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][ebh_al_subintf]['cid'] = ciq_db[node2]['Interfaces'][hub_bl_intf]['cid'] = f'''{row[h_list.index('Provider')].value}_{row[h_list.index('Circuit ID')].value}'''
                        # If the EBH AL interface is 10G, then it requires a breakout.
                        if str(row[h_list.index('EBH AL Interface')].value).count('/') >= 4:
                            ciq_db[node1]['Interfaces'][ebh_al_subintf]['Breakout'] = re.sub('^.*[A-Z|a-z]+', '', row[h_list.index('EBH AL Interface')].value).split('.')[0][:-2]
                        if str(row[h_list.index('HUB BL Interface')].value).count('/') >= 4:
                            ciq_db[node2]['Interfaces'][hub_bl_intf]['Breakout'] = re.sub('^.*[A-Z|a-z]+', '', row[h_list.index('HUB BL Interface')].value).split('.')[0][:-2]
                        # Add legacy info for HUB BL
                        if "n/a" in str(row[h_list.index('HUB BL Legacy Uplink - Far End Interface')].value).lower() or "na" in str(row[h_list.index('HUB BL Legacy Uplink - Far End Interface')].value).lower():
                            legacy_far_intf = f'''{str(row[h_list.index('HUB BL Legacy Uplink - Far End Interface')].value)}'''
                        else:
                            legacy_far_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB BL Legacy Uplink - Far End Interface')].value)).group(1)
                            legacy_far_intf_word_full = ""
                            if "gi" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "TenGigE"
                            elif "tw" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "HundredGigE"
                            legacy_far_intf = f'''{str(row[h_list.index('HUB BL Legacy Uplink - Far End Interface')].value).replace(legacy_far_intf_word, legacy_far_intf_word_full, 1)}'''
                        if "n/a" in str(row[h_list.index('HUB BL Legacy Uplink - Local Interface')].value).lower() or "na" in str(row[h_list.index('HUB BL Legacy Uplink - Local Interface')].value).lower():
                            legacy_local_intf = f'''{str(row[h_list.index('HUB BL Legacy Uplink - Local Interface')].value)}'''
                        else:
                            legacy_local_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB BL Legacy Uplink - Local Interface')].value)).group(1)
                            legacy_local_intf_word_full = ""
                            if "gi" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "TenGigE"
                            elif "tw" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "HundredGigE"
                            legacy_local_intf = f'''{str(row[h_list.index('HUB BL Legacy Uplink - Local Interface')].value).replace(legacy_local_intf_word, legacy_local_intf_word_full, 1)}'''
                        ciq_db[node2]['Interfaces'][hub_bl_intf]["HUB BL Legacy Uplink - Far End Hostname"] = f'''{row[h_list.index('HUB BL Legacy Uplink - Far End Hostname')].value}'''
                        ciq_db[node2]['Interfaces'][hub_bl_intf]["HUB BL Legacy Uplink - Far End Interface"] = f'''{legacy_far_intf}'''
                        ciq_db[node2]['Interfaces'][hub_bl_intf]["HUB BL Legacy Uplink - Local Interface"] = f'''{legacy_local_intf}'''

                        # Create a BGP dict within the device entry in ciq_db
                        if 'BGP_Peers' not in ciq_db[node1]:
                            ciq_db[node1]['BGP_Peers'] = {}
                        if 'BGP_Peers' not in ciq_db[node2]:
                            ciq_db[node2]['BGP_Peers'] = {}
                        # Create a BGP peer entry using the peer's hostname.
                        ciq_db[node1]['BGP_Peers'][node2] = {}
                        ciq_db[node2]['BGP_Peers'][node1] = {}
                        # Add the peer's device role to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node2]['eNSESR Role'] = ciq_db[node2]['eNSESR Role']
                        ciq_db[node2]['BGP_Peers'][node1]['eNSESR Role'] = ciq_db[node1]['eNSESR Role']
                        # Add the peer's Loopback 0 IP to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node2]['Loopback 0 Global - IPv4'] = ciq_db[node2]['Loopback 0 Global - IPv4']
                        ciq_db[node2]['BGP_Peers'][node1]['Loopback 0 Global - IPv4'] = ciq_db[node1]['Loopback 0 Global - IPv4']
                logger.info(f"Reading {ws} - Completed")

            # Record data provided in the sheet and add it to ciq_db.
            case 'HUB BL to HUB BL P2P':
                ws = wb[sheet]
                logger.info(f"Reading {ws} - Started")

                # Record the headers actually configured in the sheet from the CIQ.
                headers = req_headers[sheet]
                h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]

                # Go through each row and record data.
                for row in ws.iter_rows():
                    # If the 'HUB BL Hostname' is blank, stop processing data.
                    if row[h_list.index('HUB BL Hostname A')].value is None:
                        break
                    # If the row has data and the row is not the headers, record the data.
                    if row[h_list.index('HUB BL Hostname A')].row > 1:
                        # node1 is HUB BL A and node2 is HUB BL Z
                        node1 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('HUB BL Hostname A')].value.strip()][0]
                        node2 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('HUB BL Hostname Z')].value.strip()][0]
                        # Create an interface dict within the device entry in ciq_db
                        if 'Interfaces' not in ciq_db[node1]:
                            ciq_db[node1]['Interfaces'] = {}
                        if 'Interfaces' not in ciq_db[node2]:
                            ciq_db[node2]['Interfaces'] = {}
                        hub_bl_a_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB BL Interface A')].value)).group(1)
                        hub_bl_a_intf_word_full = ""
                        if "gi" in hub_bl_a_intf_word.lower()[:2]:
                            hub_bl_a_intf_word_full = "GigabitEthernet"
                        elif "te" in hub_bl_a_intf_word.lower()[:2]:
                            hub_bl_a_intf_word_full = "TenGigE"
                        elif "tw" in hub_bl_a_intf_word.lower()[:2]:
                            hub_bl_a_intf_word_full = "TwentyfiveGig"
                        elif "fo" in hub_bl_a_intf_word.lower()[:2]:
                            hub_bl_a_intf_word_full = "FortyGigE"
                        elif "hu" in hub_bl_a_intf_word.lower()[:2]:
                            hub_bl_a_intf_word_full = "HundredGigE"
                        hub_bl_z_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB BL Interface Z')].value)).group(1)
                        hub_bl_z_intf_word_full = ""
                        if "gi" in hub_bl_z_intf_word.lower()[:2]:
                            hub_bl_z_intf_word_full = "GigabitEthernet"
                        elif "te" in hub_bl_z_intf_word.lower()[:2]:
                            hub_bl_z_intf_word_full = "TenGigE"
                        elif "tw" in hub_bl_z_intf_word.lower()[:2]:
                            hub_bl_z_intf_word_full = "TwentyfiveGig"
                        elif "fo" in hub_bl_z_intf_word.lower()[:2]:
                            hub_bl_z_intf_word_full = "FortyGigE"
                        elif "hu" in hub_bl_z_intf_word.lower()[:2]:
                            hub_bl_z_intf_word_full = "HundredGigE"
                        hub_bl_a_intf = f'''{str(row[h_list.index('HUB BL Interface A')].value).replace(hub_bl_a_intf_word, hub_bl_a_intf_word_full, 1)}'''
                        hub_bl_z_intf = f'''{str(row[h_list.index('HUB BL Interface Z')].value).replace(hub_bl_z_intf_word, hub_bl_z_intf_word_full, 1)}'''
                        ciq_db[node1]['Interfaces'][hub_bl_a_intf] = {}
                        ciq_db[node2]['Interfaces'][hub_bl_z_intf] = {}
                        # Add the IPs under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_bl_a_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 A')].value}'''
                        ciq_db[node2]['Interfaces'][hub_bl_z_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 Z')].value}'''
                        # Add the interface descriptions under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_bl_a_intf]['description'] = f'''{row[h_list.index('HUB BL Hostname Z')].value}_{row[h_list.index('HUB BL Interface Z')].value}'''
                        ciq_db[node2]['Interfaces'][hub_bl_z_intf]['description'] = f'''{row[h_list.index('HUB BL Hostname A')].value}_{row[h_list.index('HUB BL Interface A')].value}'''
                        # Add the circuit ID under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_bl_a_intf]['cid'] = ciq_db[node1]['Interfaces'][hub_bl_a_intf]['description']
                        ciq_db[node2]['Interfaces'][hub_bl_z_intf]['cid'] = ciq_db[node2]['Interfaces'][hub_bl_z_intf]['description']
                        # If the interface is 10G, then it requires a breakout.
                        if str(row[h_list.index('HUB BL Interface A')].value).count('/') >= 4:
                            ciq_db[node1]['Interfaces'][hub_bl_a_intf]['Breakout'] = re.sub('^.*[A-Z|a-z]+', '', row[h_list.index('HUB BL Interface A')].value)[:-2]
                        if str(row[h_list.index('HUB BL Interface Z')].value).count('/') >= 4:
                            ciq_db[node2]['Interfaces'][hub_bl_z_intf]['Breakout'] = re.sub('^.*[A-Z|a-z]+', '', row[h_list.index('HUB BL Interface Z')].value)[:-2]

                        # Create a BGP dict within the device entry in ciq_db
                        if 'BGP_Peers' not in ciq_db[node1]:
                            ciq_db[node1]['BGP_Peers'] = {}
                        if 'BGP_Peers' not in ciq_db[node2]:
                            ciq_db[node2]['BGP_Peers'] = {}
                        # Create a BGP peer entry using the peer's hostname.
                        ciq_db[node1]['BGP_Peers'][node2] = {}
                        ciq_db[node2]['BGP_Peers'][node1] = {}
                        # Add the peer's device role to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node2]['eNSESR Role'] = ciq_db[node2]['eNSESR Role']
                        ciq_db[node2]['BGP_Peers'][node1]['eNSESR Role'] = ciq_db[node1]['eNSESR Role']
                        # Add the peer's Loopback 0 IP to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node2]['Loopback 0 Global - IPv4'] = ciq_db[node2]['Loopback 0 Global - IPv4']
                        ciq_db[node2]['BGP_Peers'][node1]['Loopback 0 Global - IPv4'] = ciq_db[node1]['Loopback 0 Global - IPv4']
                logger.info(f"Reading {ws} - Completed")

            # Record data provided in the sheet and add it to ciq_db.
            case 'HUB BL to HUB AL P2P':
                ws = wb[sheet]
                logger.info(f"Reading {ws} - Started")

                # Record the headers actually configured in the sheet from the CIQ.
                headers = req_headers[sheet]
                h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]

                # Go through each row and record data.
                for row in ws.iter_rows():
                    # If the 'HUB BL Hostname' is blank, stop processing data.
                    if row[h_list.index('HUB BL Hostname')].value is None:
                        break
                    # If the row has data and the row is not the headers, record the data.
                    if row[h_list.index('HUB AL Hostname')].row > 1:
                        # node1 is HUB BL and node2 is HUB AL
                        node1 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('HUB BL Hostname')].value.strip()][0]
                        node2 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('HUB AL Hostname')].value.strip()][0]
                        # Create an interface dict within the device entry in ciq_db
                        if 'Interfaces' not in ciq_db[node1]:
                            ciq_db[node1]['Interfaces'] = {}
                        if 'Interfaces' not in ciq_db[node2]:
                            ciq_db[node2]['Interfaces'] = {}
                        hub_bl_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB BL Interface')].value)).group(1)
                        hub_bl_intf_word_full = ""
                        if "gi" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "GigabitEthernet"
                        elif "te" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "TenGigE"
                        elif "tw" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "TwentyfiveGig"
                        elif "fo" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "FortyGigE"
                        elif "hu" in hub_bl_intf_word.lower()[:2]:
                            hub_bl_intf_word_full = "HundredGigE"
                        hub_al_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB AL Interface')].value)).group(1)
                        hub_al_intf_word_full = ""
                        if "gi" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "GigabitEthernet"
                        elif "te" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "TenGigE"
                        elif "tw" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "TwentyfiveGig"
                        elif "fo" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "FortyGigE"
                        elif "hu" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "HundredGigE"
                        hub_bl_intf = f'''{str(row[h_list.index('HUB BL Interface')].value).replace(hub_bl_intf_word, hub_bl_intf_word_full, 1)}'''
                        hub_al_intf = f'''{str(row[h_list.index('HUB AL Interface')].value).replace(hub_al_intf_word, hub_al_intf_word_full, 1)}'''
                        ciq_db[node1]['Interfaces'][hub_bl_intf] = {}
                        ciq_db[node2]['Interfaces'][hub_al_intf] = {}
                        # Add the IPs under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_bl_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 A')].value}'''
                        ciq_db[node2]['Interfaces'][hub_al_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 Z')].value}'''
                        # Add the interface descriptions under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_bl_intf]['description'] = f'''{row[h_list.index('HUB AL Hostname')].value}_{row[h_list.index('HUB AL Interface')].value}'''
                        ciq_db[node2]['Interfaces'][hub_al_intf]['description'] = f'''{row[h_list.index('HUB BL Hostname')].value}_{row[h_list.index('HUB BL Interface')].value}'''
                        # Add the circuit ID under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_bl_intf]['cid'] = ciq_db[node1]['Interfaces'][hub_bl_intf]['description']
                        ciq_db[node2]['Interfaces'][hub_al_intf]['cid'] = ciq_db[node2]['Interfaces'][hub_al_intf]['description']
                        # If the interface is 10G, then it requires a breakout.
                        if str(row[h_list.index('HUB BL Interface')].value).count('/') >= 4:
                            ciq_db[node1]['Interfaces'][hub_bl_intf]['Breakout'] = re.sub('^.*[A-Z|a-z]+', '', row[h_list.index('HUB BL Interface')].value)[:-2]
                        if str(row[h_list.index('HUB AL Interface')].value).count('/') >= 4:
                            ciq_db[node2]['Interfaces'][hub_al_intf]['Breakout'] = re.sub('^.*[A-Z|a-z]+', '', row[h_list.index('HUB AL Interface')].value)[:-2]
                        # Add legacy info for HUB AL
                        if "n/a" in str(row[h_list.index('HUB AL Legacy Uplink - Far End Interface')].value).lower() or "na" in str(row[h_list.index('HUB AL Legacy Uplink - Far End Interface')].value).lower():
                            legacy_far_intf = f'''{str(row[h_list.index('HUB AL Legacy Uplink - Far End Interface')].value)}'''
                        else:
                            legacy_far_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB AL Legacy Uplink - Far End Interface')].value)).group(1)
                            legacy_far_intf_word_full = ""
                            if "gi" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "TenGigE"
                            elif "tw" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "HundredGigE"
                            legacy_far_intf = f'''{str(row[h_list.index('HUB AL Legacy Uplink - Far End Interface')].value).replace(legacy_far_intf_word, legacy_far_intf_word_full, 1)}'''
                        if "n/a" in str(row[h_list.index('HUB AL Legacy Uplink - Local Interface')].value).lower() or "na" in str(row[h_list.index('HUB AL Legacy Uplink - Local Interface')].value).lower():
                            legacy_local_intf = f'''{str(row[h_list.index('HUB AL Legacy Uplink - Local Interface')].value)}'''
                        else:
                            legacy_local_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB AL Legacy Uplink - Local Interface')].value)).group(1)
                            legacy_local_intf_word_full = ""
                            if "gi" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "TenGigE"
                            elif "tw" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "HundredGigE"
                            legacy_local_intf = f'''{str(row[h_list.index('HUB AL Legacy Uplink - Local Interface')].value).replace(legacy_local_intf_word, legacy_local_intf_word_full, 1)}'''
                        ciq_db[node2]['Interfaces'][hub_al_intf]["HUB AL Legacy Uplink - Far End Hostname"] = f'''{row[h_list.index('HUB AL Legacy Uplink - Far End Hostname')].value}'''
                        ciq_db[node2]['Interfaces'][hub_al_intf]["HUB AL Legacy Uplink - Far End Interface"] = f'''{legacy_far_intf}'''
                        ciq_db[node2]['Interfaces'][hub_al_intf]["HUB AL Legacy Uplink - Local Interface"] = f'''{legacy_local_intf}'''

                        # Create a BGP dict within the device entry in ciq_db
                        if 'BGP_Peers' not in ciq_db[node1]:
                            ciq_db[node1]['BGP_Peers'] = {}
                        if 'BGP_Peers' not in ciq_db[node2]:
                            ciq_db[node2]['BGP_Peers'] = {}
                        # Create a BGP peer entry using the peer's hostname.
                        ciq_db[node1]['BGP_Peers'][node2] = {}
                        ciq_db[node2]['BGP_Peers'][node1] = {}
                        # Add the peer's device role to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node2]['eNSESR Role'] = ciq_db[node2]['eNSESR Role']
                        ciq_db[node2]['BGP_Peers'][node1]['eNSESR Role'] = ciq_db[node1]['eNSESR Role']
                        # Add the peer's Loopback 0 IP to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node2]['Loopback 0 Global - IPv4'] = ciq_db[node2]['Loopback 0 Global - IPv4']
                        ciq_db[node2]['BGP_Peers'][node1]['Loopback 0 Global - IPv4'] = ciq_db[node1]['Loopback 0 Global - IPv4']
                logger.info(f"Reading {ws} - Completed")

            # Record data provided in the sheet and add it to ciq_db.
            case 'HUB AL to SPOKE P2P':
                ws = wb[sheet]
                logger.info(f"Reading {ws} - Started")

                # Record the headers actually configured in the sheet from the CIQ.
                headers = req_headers[sheet]
                h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]

                # Go through each row and record data.
                for row in ws.iter_rows():
                    # If the 'HUB AL Hostname' is blank, stop processing data.
                    if row[h_list.index('HUB AL Hostname')].value is None:
                        break
                    # If the row has data and the row is not the headers, record the data.
                    if row[h_list.index('HUB AL Hostname')].row > 1:
                        # node1 is the HUB and node2 is the Spoke
                        node1 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('HUB AL Hostname')].value.strip()][0]
                        node2 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('SPOKE Hostname')].value.strip()][0]
                        # Create an interface dict within the device entry in ciq_db
                        if 'Interfaces' not in ciq_db[node1]:
                            ciq_db[node1]['Interfaces'] = {}
                        if 'Interfaces' not in ciq_db[node2]:
                            ciq_db[node2]['Interfaces'] = {}
                        hub_al_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB AL Interface')].value)).group(1)
                        hub_al_intf_word_full = ""
                        if "gi" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "GigabitEthernet"
                        elif "te" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "TenGigE"
                        elif "tw" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "TwentyfiveGig"
                        elif "fo" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "FortyGigE"
                        elif "hu" in hub_al_intf_word.lower()[:2]:
                            hub_al_intf_word_full = "HundredGigE"
                        spoke_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('SPOKE Interface')].value)).group(1)
                        spoke_intf_word_full = ""
                        if "gi" in spoke_intf_word.lower()[:2]:
                            spoke_intf_word_full = "GigabitEthernet"
                        elif "te" in spoke_intf_word.lower()[:2]:
                            spoke_intf_word_full = "TenGigE"
                        elif "tw" in spoke_intf_word.lower()[:2]:
                            spoke_intf_word_full = "TwentyfiveGig"
                        elif "fo" in spoke_intf_word.lower()[:2]:
                            spoke_intf_word_full = "FortyGigE"
                        elif "hu" in spoke_intf_word.lower()[:2]:
                            spoke_intf_word_full = "HundredGigE"
                        hub_al_intf = f'''{str(row[h_list.index('HUB AL Interface')].value).replace(hub_al_intf_word, hub_al_intf_word_full, 1)}.{row[h_list.index('VLAN')].value}'''
                        spoke_intf = f'''{str(row[h_list.index('SPOKE Interface')].value).replace(spoke_intf_word, spoke_intf_word_full, 1)}.{row[h_list.index('VLAN')].value}'''
                        ciq_db[node1]['Interfaces'][hub_al_intf] = {}
                        ciq_db[node2]['Interfaces'][spoke_intf] = {}
                        # Add the VLAN
                        ciq_db[node1]['Interfaces'][hub_al_intf]['VLAN'] = f'''{row[h_list.index('VLAN')].value}'''
                        ciq_db[node2]['Interfaces'][spoke_intf]['VLAN'] = f'''{row[h_list.index('VLAN')].value}'''
                        # Add the IPs under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_al_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 A')].value}'''
                        ciq_db[node2]['Interfaces'][spoke_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 Z')].value}'''
                        # Add the interface descriptions under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_al_intf]['description'] = f'''{row[h_list.index('SPOKE Hostname')].value}_{row[h_list.index('SPOKE Interface')].value}'''
                        ciq_db[node2]['Interfaces'][spoke_intf]['description'] = f'''{row[h_list.index('HUB AL Hostname')].value}_{row[h_list.index('HUB AL Interface')].value}'''
                        # Add the circuit rate under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_al_intf]['circuit_rate'] = ciq_db[node2]['Interfaces'][spoke_intf]['circuit_rate'] = f'''{row[h_list.index('Circuit Rate (Mbps)')].value}'''
                        # Add the circuit ID under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][hub_al_intf]['cid'] = ciq_db[node2]['Interfaces'][spoke_intf]['cid'] = f'''{row[h_list.index('Provider')].value}_{row[h_list.index('Circuit ID')].value}'''
                        # Add legacy info for SPOKE
                        if "n/a" in str(row[h_list.index('SPOKE Legacy Uplink - Far End Interface')].value).lower() or "na" in str(row[h_list.index('SPOKE Legacy Uplink - Far End Interface')].value).lower():
                            legacy_far_intf = f'''{str(row[h_list.index('SPOKE Legacy Uplink - Far End Interface')].value)}'''
                        else:
                            legacy_far_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('SPOKE Legacy Uplink - Far End Interface')].value)).group(1)
                            legacy_far_intf_word_full = ""
                            if "gi" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "TenGigE"
                            elif "tw" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "HundredGigE"
                            legacy_far_intf = f'''{str(row[h_list.index('SPOKE Legacy Uplink - Far End Interface')].value).replace(legacy_far_intf_word, legacy_far_intf_word_full, 1)}'''
                        if "n/a" in str(row[h_list.index('SPOKE Legacy Uplink - Local Interface')].value).lower() or "na" in str(row[h_list.index('SPOKE Legacy Uplink - Local Interface')].value).lower():
                            legacy_local_intf = f'''{str(row[h_list.index('SPOKE Legacy Uplink - Local Interface')].value)}'''
                        else:
                            legacy_local_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('SPOKE Legacy Uplink - Local Interface')].value)).group(1)
                            legacy_local_intf_word_full = ""
                            if "gi" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "TenGigE"
                            elif "tw" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "HundredGigE"
                            legacy_local_intf = f'''{str(row[h_list.index('SPOKE Legacy Uplink - Local Interface')].value).replace(legacy_local_intf_word, legacy_local_intf_word_full, 1)}'''
                        ciq_db[node2]['Interfaces'][spoke_intf]["SPOKE Legacy Uplink - Far End Hostname"] = f'''{row[h_list.index('SPOKE Legacy Uplink - Far End Hostname')].value}'''
                        ciq_db[node2]['Interfaces'][spoke_intf]["SPOKE Legacy Uplink - Far End Interface"] = f'''{legacy_far_intf}'''
                        ciq_db[node2]['Interfaces'][spoke_intf]["SPOKE Legacy Uplink - Local Interface"] = f'''{legacy_local_intf}'''

                        # Create a BGP dict within the device entry in ciq_db
                        if 'BGP_Peers' not in ciq_db[node1]:
                            ciq_db[node1]['BGP_Peers'] = {}
                        if 'BGP_Peers' not in ciq_db[node2]:
                            ciq_db[node2]['BGP_Peers'] = {}
                        # Create a BGP peer entry using the peer's hostname.
                        ciq_db[node1]['BGP_Peers'][node2] = {}
                        ciq_db[node2]['BGP_Peers'][node1] = {}
                        # Add the peer's device role to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node2]['eNSESR Role'] = ciq_db[node2]['eNSESR Role']
                        ciq_db[node2]['BGP_Peers'][node1]['eNSESR Role'] = ciq_db[node1]['eNSESR Role']
                        # Add the peer's Loopback 0 IP to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node2]['Loopback 0 Global - IPv4'] = ciq_db[node2]['Loopback 0 Global - IPv4']
                        ciq_db[node2]['BGP_Peers'][node1]['Loopback 0 Global - IPv4'] = ciq_db[node1]['Loopback 0 Global - IPv4']
                logger.info(f"Reading {ws} - Completed")

            # Record data provided in the sheet and add it to ciq_db.
            case 'SPOKE to SPOKE P2P':
                ws = wb[sheet]
                logger.info(f"Reading {ws} - Started")

                # Record the headers actually configured in the sheet from the CIQ.
                headers = req_headers[sheet]
                h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]

                # Go through each row and record data.
                for row in ws.iter_rows():
                    # If the 'Upstream SPOKE Hostname Connected to HUB AL' is blank, stop processing data.
                    if row[h_list.index('Upstream SPOKE Hostname Connected to HUB AL')].value is None:
                        break
                    # If the row has data and the row is not the headers, record the data.
                    if row[h_list.index('Upstream SPOKE Hostname Connected to HUB AL')].row > 1:
                        # node1 is the Upstream SPOKE, node2 is the Downstream SPOKE, node3 is the Upstream HUB AL
                        node1 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('Upstream SPOKE Hostname')].value.strip()][0]
                        node2 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('Downstream SPOKE Hostname')].value.strip()][0]
                        node3 = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('Upstream SPOKE Hostname Connected to HUB AL')].value.strip()][0]
                        # Create an interface dict within the device entry in ciq_db
                        if 'Interfaces' not in ciq_db[node1]:
                            ciq_db[node1]['Interfaces'] = {}
                        if 'Interfaces' not in ciq_db[node2]:
                            ciq_db[node2]['Interfaces'] = {}
                        up_spoke_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('Upstream SPOKE Interface')].value)).group(1)
                        up_spoke_intf_word_full = ""
                        if "gi" in up_spoke_intf_word.lower()[:2]:
                            up_spoke_intf_word_full = "GigabitEthernet"
                        elif "te" in up_spoke_intf_word.lower()[:2]:
                            up_spoke_intf_word_full = "TenGigE"
                        elif "tw" in up_spoke_intf_word.lower()[:2]:
                            up_spoke_intf_word_full = "TwentyfiveGig"
                        elif "fo" in up_spoke_intf_word.lower()[:2]:
                            up_spoke_intf_word_full = "FortyGigE"
                        elif "hu" in up_spoke_intf_word.lower()[:2]:
                            up_spoke_intf_word_full = "HundredGigE"
                        down_spoke_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('Downstream SPOKE Interface')].value)).group(1)
                        down_spoke_intf_word_full = ""
                        if "gi" in down_spoke_intf_word.lower()[:2]:
                            down_spoke_intf_word_full = "GigabitEthernet"
                        elif "te" in down_spoke_intf_word.lower()[:2]:
                            down_spoke_intf_word_full = "TenGigE"
                        elif "tw" in down_spoke_intf_word.lower()[:2]:
                            down_spoke_intf_word_full = "TwentyfiveGig"
                        elif "fo" in down_spoke_intf_word.lower()[:2]:
                            down_spoke_intf_word_full = "FortyGigE"
                        elif "hu" in down_spoke_intf_word.lower()[:2]:
                            down_spoke_intf_word_full = "HundredGigE"
                        up_spoke_intf = f'''{str(row[h_list.index('Upstream SPOKE Interface')].value).replace(up_spoke_intf_word, up_spoke_intf_word_full, 1)}.{row[h_list.index('VLAN')].value}'''
                        down_spoke_intf = f'''{str(row[h_list.index('Downstream SPOKE Interface')].value).replace(down_spoke_intf_word, down_spoke_intf_word_full, 1)}.{row[h_list.index('VLAN')].value}'''
                        ciq_db[node1]['Interfaces'][up_spoke_intf] = {}
                        ciq_db[node2]['Interfaces'][down_spoke_intf] = {}
                        # Add the VLAN
                        ciq_db[node1]['Interfaces'][up_spoke_intf]['VLAN'] = f'''{row[h_list.index('VLAN')].value}'''
                        ciq_db[node2]['Interfaces'][down_spoke_intf]['VLAN'] = f'''{row[h_list.index('VLAN')].value}'''
                        # Add the IPs under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][up_spoke_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 A')].value}'''
                        ciq_db[node2]['Interfaces'][down_spoke_intf]['ip'] = f'''{row[h_list.index('IPv4 /31 Z')].value}'''
                        # Add the interface descriptions under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][up_spoke_intf]['description'] = f'''{row[h_list.index('Downstream SPOKE Hostname')].value}_{row[h_list.index('Downstream SPOKE Interface')].value}'''
                        ciq_db[node2]['Interfaces'][down_spoke_intf]['description'] = f'''{row[h_list.index('Upstream SPOKE Hostname')].value}_{row[h_list.index('Upstream SPOKE Interface')].value}'''
                        # Add the circuit rate under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][up_spoke_intf]['circuit_rate'] = ciq_db[node2]['Interfaces'][down_spoke_intf]['circuit_rate'] = f'''{row[h_list.index('Circuit Rate (Mbps)')].value}'''
                        # Add the circuit ID under each interface within the device entry in ciq_db
                        ciq_db[node1]['Interfaces'][up_spoke_intf]['cid'] = ciq_db[node2]['Interfaces'][down_spoke_intf]['cid'] = f'''{row[h_list.index('Provider')].value}_{row[h_list.index('Circuit ID')].value}'''
                        # Add legacy info for Downstream SPOKE
                        if "n/a" in str(row[h_list.index('Downstream SPOKE Legacy Uplink - Far End Interface')].value).lower() or "na" in str(row[h_list.index('Downstream SPOKE Legacy Uplink - Far End Interface')].value).lower():
                            legacy_far_intf = f'''{str(row[h_list.index('Downstream SPOKE Legacy Uplink - Far End Interface')].value)}'''
                        else:
                            legacy_far_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('Downstream SPOKE Legacy Uplink - Far End Interface')].value)).group(1)
                            legacy_far_intf_word_full = ""
                            if "gi" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "TenGigE"
                            elif "tw" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_far_intf_word.lower()[:2]:
                                legacy_far_intf_word_full = "HundredGigE"
                            legacy_far_intf = f'''{str(row[h_list.index('Downstream SPOKE Legacy Uplink - Far End Interface')].value).replace(legacy_far_intf_word, legacy_far_intf_word_full, 1)}'''
                        if "n/a" in str(row[h_list.index('Downstream SPOKE Legacy Uplink - Local Interface')].value).lower() or "na" in str(row[h_list.index('Downstream SPOKE Legacy Uplink - Local Interface')].value).lower():
                            legacy_local_intf = f'''{str(row[h_list.index('Downstream SPOKE Legacy Uplink - Local Interface')].value)}'''
                        else:
                            legacy_local_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('Downstream SPOKE Legacy Uplink - Local Interface')].value)).group(1)
                            legacy_local_intf_word_full = ""
                            if "gi" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "TenGigE"
                            elif "tw" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_local_intf_word.lower()[:2]:
                                legacy_local_intf_word_full = "HundredGigE"
                            legacy_local_intf = f'''{str(row[h_list.index('Downstream SPOKE Legacy Uplink - Local Interface')].value).replace(legacy_local_intf_word, legacy_local_intf_word_full, 1)}'''
                        ciq_db[node2]['Interfaces'][down_spoke_intf]["Downstream SPOKE Legacy Uplink - Far End Hostname"] = f'''{row[h_list.index('Downstream SPOKE Legacy Uplink - Far End Hostname')].value}'''
                        ciq_db[node2]['Interfaces'][down_spoke_intf]["Downstream SPOKE Legacy Uplink - Far End Interface"] = f'''{legacy_far_intf}'''
                        ciq_db[node2]['Interfaces'][down_spoke_intf]["Downstream SPOKE Legacy Uplink - Local Interface"] = f'''{legacy_local_intf}'''

                        # Create a BGP dict within the device entry in ciq_db
                        if 'BGP_Peers' not in ciq_db[node1]:
                            ciq_db[node1]['BGP_Peers'] = {}
                        if 'BGP_Peers' not in ciq_db[node2]:
                            ciq_db[node2]['BGP_Peers'] = {}
                        if 'BGP_Peers' not in ciq_db[node3]:
                            ciq_db[node3]['BGP_Peers'] = {}
                        # Create a BGP peer entry using the peer's hostname. Both spokes will peer with the HUB AL
                        ciq_db[node1]['BGP_Peers'][node3] = {}
                        ciq_db[node2]['BGP_Peers'][node3] = {}
                        ciq_db[node3]['BGP_Peers'][node1] = {}
                        ciq_db[node3]['BGP_Peers'][node2] = {}
                        # Add the peer's device role to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node3]['eNSESR Role'] = ciq_db[node3]['eNSESR Role']
                        ciq_db[node2]['BGP_Peers'][node3]['eNSESR Role'] = ciq_db[node3]['eNSESR Role']
                        ciq_db[node3]['BGP_Peers'][node1]['eNSESR Role'] = ciq_db[node1]['eNSESR Role']
                        ciq_db[node3]['BGP_Peers'][node2]['eNSESR Role'] = ciq_db[node2]['eNSESR Role']
                        # Add the peer's Loopback 0 IP to the peer entry.
                        ciq_db[node1]['BGP_Peers'][node3]['Loopback 0 Global - IPv4'] = ciq_db[node3]['Loopback 0 Global - IPv4']
                        ciq_db[node2]['BGP_Peers'][node3]['Loopback 0 Global - IPv4'] = ciq_db[node3]['Loopback 0 Global - IPv4']
                        ciq_db[node3]['BGP_Peers'][node1]['Loopback 0 Global - IPv4'] = ciq_db[node1]['Loopback 0 Global - IPv4']
                        ciq_db[node3]['BGP_Peers'][node2]['Loopback 0 Global - IPv4'] = ciq_db[node2]['Loopback 0 Global - IPv4']

                        # If this is the HUB DRAN Spoke, remove the BGP peer to itself that was added before.
                        if node1 == node3:
                            del ciq_db[node1]['BGP_Peers'][node3]
                logger.info(f"Reading {ws} - Completed")

            # Record data provided in the sheet and add it to ciq_db.
            case 'Non-RAN Services':
                ws = wb[sheet]
                logger.info(f"Reading {ws} - Started")

                # Record the headers actually configured in the sheet from the CIQ.
                headers = req_headers[sheet]
                h_list = [x[0].value for x in ws.iter_cols(min_col=None, max_col=ws.max_column, min_row=1, max_row=1)]
                edn_host = ""
                # Go through each row and record data.
                for row in ws.iter_rows():
                    # If the 'Hostname' is blank, stop processing data.
                    if row[h_list.index('HUB Device Hostname')].value is None:
                        break
                    # If the row has data and the row is not the headers, record the data.
                    if row[h_list.index('HUB Device Hostname')].row > 1:
                        if row[h_list.index('Non-RAN Device Hostname')].value not in ciq_db:
                            ciq_db[row[h_list.index('Non-RAN Device Hostname')].value] = {}
                            ciq_db[row[h_list.index('Non-RAN Device Hostname')].value]['New Hostname'] = row[h_list.index('Non-RAN Device Hostname')].value
                        node = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('HUB Device Hostname')].value][0]
                        non_ran_node = [x for x in ciq_db if ciq_db[x]['New Hostname'].strip() == row[h_list.index('Non-RAN Device Hostname')].value][0]
                        ciq_db[non_ran_node]['eNSESR Role'] = "NON RAN"
                        # Create a Non-RAN dict within the device entry in ciq_db
                        if 'Non-RAN' not in ciq_db[node]:
                            ciq_db[node]['Non-RAN'] = {}
                        if 'Non-RAN' not in ciq_db[non_ran_node]:
                            ciq_db[non_ran_node]['Non-RAN'] = {}
                        node_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('HUB Device Interface')].value)).group(1)
                        node_intf_word_full = ""
                        if "gi" in node_intf_word.lower()[:2]:
                            node_intf_word_full = "GigabitEthernet"
                        elif "te" in node_intf_word.lower()[:2]:
                            node_intf_word_full = "TenGigE"
                        elif "tw" in node_intf_word.lower()[:2]:
                            node_intf_word_full = "TwentyfiveGig"
                        elif "fo" in node_intf_word.lower()[:2]:
                            node_intf_word_full = "FortyGigE"
                        elif "hu" in node_intf_word.lower()[:2]:
                            node_intf_word_full = "HundredGigE"
                        node_intf = f'''{str(row[h_list.index('HUB Device Interface')].value).replace(node_intf_word, node_intf_word_full, 1)}'''
                        non_ran_node_intf = f'''{str(row[h_list.index('Non-RAN Device Interface')].value)}'''
                        ciq_db[node]['Non-RAN'][node_intf] = {}
                        ciq_db[non_ran_node]['Non-RAN'][non_ran_node_intf] = {}
                        ciq_db[node]['Non-RAN'][node_intf]['HUB Device Interface'] = f'''{node_intf}'''
                        ciq_db[non_ran_node]['Non-RAN'][non_ran_node_intf]['Non-RAN Device Interface'] = f'''{non_ran_node_intf}'''
                        ciq_db[node]['Non-RAN'][node_intf]['Type'] = ciq_db[non_ran_node]['Non-RAN'][non_ran_node_intf]['Type'] = f'''{row[h_list.index('Type')].value}'''
                        if "EDN" in ciq_db[node]['Non-RAN'][node_intf]['Type']:
                            edn_host = ciq_db[node]["New Hostname"]
                        ciq_db[node]['Non-RAN'][node_intf]['Non-RAN Device Hostname'] = f'''{row[h_list.index('Non-RAN Device Hostname')].value}'''
                        ciq_db[non_ran_node]['Non-RAN'][non_ran_node_intf]['HUB Device Hostname'] = f'''{row[h_list.index('HUB Device Hostname')].value}'''
                        ciq_db[node]['Non-RAN'][node_intf]['Non-RAN Device Interface'] = f'''{non_ran_node_intf}'''
                        ciq_db[non_ran_node]['Non-RAN'][non_ran_node_intf]['HUB Device Interface'] = f'''{node_intf}'''
                        # Add legacy info for Non-RAN Device
                        if "n/a" in str(row[h_list.index('Non-RAN Device Legacy Uplink - Far End Interface')].value).lower() or "na" in str(row[h_list.index('Non-RAN Device Legacy Uplink - Far End Interface')].value).lower():
                            legacy_non_ran_node_far_intf = f'''{str(row[h_list.index('Non-RAN Device Legacy Uplink - Far End Interface')].value)}'''
                        else:
                            legacy_non_ran_node_far_intf_word = re.search(r"([a-z|A-Z]+)\d.*", str(row[h_list.index('Non-RAN Device Legacy Uplink - Far End Interface')].value)).group(1)
                            legacy_non_ran_node_far_intf_word_full = ""
                            if "gi" in legacy_non_ran_node_far_intf_word.lower()[:2]:
                                legacy_non_ran_node_far_intf_word_full = "GigabitEthernet"
                            elif "te" in legacy_non_ran_node_far_intf_word.lower()[:2]:
                                legacy_non_ran_node_far_intf_word_full = "TenGigE"
                            elif "tw" in legacy_non_ran_node_far_intf_word.lower()[:2]:
                                legacy_non_ran_node_far_intf_word_full = "TwentyfiveGig"
                            elif "fo" in legacy_non_ran_node_far_intf_word.lower()[:2]:
                                legacy_non_ran_node_far_intf_word_full = "FortyGigE"
                            elif "hu" in legacy_non_ran_node_far_intf_word.lower()[:2]:
                                legacy_non_ran_node_far_intf_word_full = "HundredGigE"
                            legacy_non_ran_node_far_intf = f'''{str(row[h_list.index('Non-RAN Device Legacy Uplink - Far End Interface')].value).replace(legacy_non_ran_node_far_intf_word, legacy_non_ran_node_far_intf_word_full, 1)}'''
                        ciq_db[non_ran_node]['Non-RAN'][non_ran_node_intf]["Non-RAN Device Legacy Uplink - Far End Hostname"] = f'''{row[h_list.index('Non-RAN Device Legacy Uplink - Far End Hostname')].value}'''
                        ciq_db[non_ran_node]['Non-RAN'][non_ran_node_intf]["Non-RAN Device Legacy Uplink - Far End Interface"] = f'''{legacy_non_ran_node_far_intf}'''
                        ciq_db[non_ran_node]['Non-RAN'][non_ran_node_intf]["Non-RAN Device Legacy Uplink - Local Interface"] = f'''{row[h_list.index('Non-RAN Device Legacy Uplink - Local Interface')].value}'''

                        # If the interface is 10G, then it requires a breakout.
                        if str(row[h_list.index('HUB Device Interface')].value).count('/') >= 4:
                            ciq_db[node]['Non-RAN'][node_intf]['Breakout'] = re.sub('^.*[A-Z|a-z]+', '', row[h_list.index('HUB Device Interface')].value)[:-2]

                # Go through each device in the db and add the hostname of the EDN host
                for node in ciq_db:
                    ciq_db[node]['EDN-Host'] = edn_host
                    continue
                logger.info(f"Reading {ws} - Completed")

    # Update the BGP_Peers so the solution is identified for each peer (L3VPN or EVPN).
    for device in ciq_db:
        if "EBH AL" in ciq_db[device]["eNSESR Role"]:
            continue
        elif "HUB BL" in ciq_db[device]["eNSESR Role"]:
            continue
        elif "HUB AL" in ciq_db[device]["eNSESR Role"]:
            for peer in ciq_db[device]["BGP_Peers"]:
                if "-cn-" in peer.lower():
                    ciq_db[device]["BGP_Peers"][peer]["Solution"] = "EVPN"
                    ciq_db[peer]["BGP_Peers"][device]["Solution"] = "EVPN"
                else:
                    ciq_db[device]["BGP_Peers"][peer]["Solution"] = "EVPN"
                    ciq_db[peer]["BGP_Peers"][device]["Solution"] = "EVPN"
        elif "DRAN SPOKE" in ciq_db[device]["eNSESR Role"]:
            if "-cn-" in device.lower():
                for peer in ciq_db[device]["BGP_Peers"]:
                    ciq_db[device]["BGP_Peers"][peer]["Solution"] = "EVPN"
                    ciq_db[peer]["BGP_Peers"][device]["Solution"] = "EVPN"
                    if "hub al" in ciq_db[peer]["eNSESR Role"].lower():
                        continue
                    else:
                        for s_peer in ciq_db[peer]["BGP_Peers"]:
                            ciq_db[peer]["BGP_Peers"][s_peer]["Solution"] = "EVPN"
                            ciq_db[s_peer]["BGP_Peers"][peer]["Solution"] = "EVPN"
            else:
                for peer in ciq_db[device]["BGP_Peers"]:
                    if "Solution" in ciq_db[device]["BGP_Peers"][peer].keys():
                        continue
                    else:
                        ciq_db[device]["BGP_Peers"][peer]["Solution"] = "EVPN"

    logger.info("Worksheet validation - Completed")
    return ciq_db, site_name


def hub_pings(ciq_db):
    mls_lte_pings = ""
    mls_cell_mgmt_pings = ""
    sr_ran_pings = ""
    sr_cell_mgmt_pings = ""
    for dev in ciq_db:
        if "B40" in dev:
            continue
        elif "NON RAN" in ciq_db[dev]["eNSESR Role"]:
            continue
        else:
            mls_lte_pings += f'''ping vrf LTE {ciq_db[dev]["Loopback 1 RAN - IPv6"]} source loopback400\n'''
            mls_cell_mgmt_pings += f'''ping vrf CELL_MGMT {ciq_db[dev]["Loopback 4 CELL_MGMT - IPv6"]} source loopback300\n'''
            sr_ran_pings += f'''ping vrf RAN {ciq_db[dev]["Loopback 1 RAN - IPv6"]} source loopback1\n'''
            sr_cell_mgmt_pings += f'''ping vrf CELL_MGMT {ciq_db[dev]["Loopback 4 CELL_MGMT - IPv6"]} source loopback4\n'''
    conf = "# Pings from the MLS\n"
    conf += mls_lte_pings
    conf += mls_cell_mgmt_pings
    conf += "\n# Pings from the eNSE SR\n"
    conf += sr_ran_pings
    conf += sr_cell_mgmt_pings
    return conf


def np_auth():
    # Instantiate a new Mimir client object
    m = Mimir()
    # Authenticate with NP using user credentials.
    try:
        m.authenticate()
    except MimirAuthenticationError as e:
        print(e)
        sys.exit(1)
    except Exception as e:
        print(e)
        sys.exit(2)
    return m


def get_np(m, host, today, last_week):
    # Network Profiler Company ID for Verizon Wireless
    company_id = "94617"

    device_details = m.np.device_details.get(cpyKey=company_id, deviceName=host)
    if device_details:
        try:
            device_details_entry = next(device_details)
            device_config_time = datetime.strptime(device_details_entry.configTime, '%Y-%m-%dT%H:%M:%S')
            device_date = device_config_time.date()
            if not (last_week <= device_date <= today):
                # Config was NOT pulled by NP within the last week.
                print(f"##### {host} config file was in NP, but it's more than 1 week old. Save the config file manually and then re-run the script.")
                return "old"
        except:
            print(f"{host} config file not found in NP. Save the config file manually and then re-run the script.")
            config_info = ""
            return config_info

    # Get the standard running config for the device from NP.
    device_np = m.np.devices.get(cpyKey=company_id, deviceName=host)
    # Check if the NP response contains a config.
    if device_np:
        try:
            device_info = next(device_np)
        except:
            config_info = ""
            return config_info
        config_np = m.np.config.get(cpyKey=company_id, deviceId=device_info.deviceId, configType='STANDARD RUNNING')
        try:
            config_info = next(config_np)
            return config_info.rawData
        except:
            config_info = ""
            return config_info


def build_topology(ciq_db, topo_directory_name, site_name):
    logger.info("Generating the Topology Diagram - Started")

    # Create an empty graph to work from
    topo = nx.Graph()

    node_ranking = ["EBH AL", "HUB BL", "PTP", "EDN", "HUB AL", "DRAN SPOKE", "NON RAN"]

    nodes_list1 = []
    nodes_list2 = []
    edge_labels = {}

    for node in ciq_db:
        if "NON RAN" in ciq_db[node]["eNSESR Role"]:
            continue
        # Add the node to the topology.
        topo.add_node(node, rank=node_ranking.index(ciq_db[node]["eNSESR Role"]))
        if "DRAN SPOKE" not in ciq_db[node]["eNSESR Role"]:
            nodes_list1.append(node)
        for intf in ciq_db[node]["Interfaces"]:
                if ciq_db[node]["Interfaces"][intf]["description"].split('_')[0] in ciq_db:
                    node2 = ciq_db[node]["Interfaces"][intf]["description"].split('_')[0]
                    # Add the link to the topology.
                    topo.add_edge(node, node2)
                    if "DRAN SPOKE" in ciq_db[node2]["eNSESR Role"]:
                        nodes_list2.append([node, node2])
                if "Non-RAN" in ciq_db[node]:
                    for intf in ciq_db[node]["Non-RAN"]:
                        if "PTP" in ciq_db[node]["Non-RAN"][intf]["Type"]:
                            ptp_hostname = ciq_db[node]["Non-RAN"][intf]["Non-RAN Device Hostname"]
                            if int(ciq_db[node]["Non-RAN"][intf]["Non-RAN Device Hostname"][-2]) % 2 == 1:
                                topo.add_node(ptp_hostname, rank=node_ranking.index("PTP"))
                                topo.add_edge(node, ptp_hostname)
                                edge_labels.update({(node, ptp_hostname): ciq_db[node]["Non-RAN"][intf]["Type"]})
                            elif int(ciq_db[node]["Non-RAN"][intf]["Non-RAN Device Hostname"][-2]) % 2 == 0:
                                topo.add_node(ptp_hostname, rank=node_ranking.index("PTP"))
                                topo.add_edge(node, ptp_hostname)
                                edge_labels.update({(node, ptp_hostname): ciq_db[node]["Non-RAN"][intf]["Type"]})
                            else:
                                logger.critical(f"Can't add PTP server <{ptp_hostname}> to Topology because last 2 chars are not digits.")
                        elif "EDN" in ciq_db[node]["Non-RAN"][intf]["Type"]:
                            edn_hostname = ciq_db[node]["Non-RAN"][intf]["Non-RAN Device Hostname"]
                            if int(ciq_db[node]["Non-RAN"][intf]["Non-RAN Device Hostname"][-2]) % 2 == 1:
                                topo.add_node(edn_hostname, rank=node_ranking.index("EDN"))
                                topo.add_edge(node, edn_hostname)
                                edge_labels.update({(node, edn_hostname): ciq_db[node]["Non-RAN"][intf]["Type"]})
                            elif int(ciq_db[node]["Non-RAN"][intf]["Non-RAN Device Hostname"][-2]) % 2 == 0:
                                topo.add_node(edn_hostname, rank=node_ranking.index("EDN"))
                                topo.add_edge(node, edn_hostname)
                                edge_labels.update({(node, edn_hostname): ciq_db[node]["Non-RAN"][intf]["Type"]})
                            else:
                                logger.critical(f"Can't add EDN switch <{edn_hostname}> to Topology because last 2 chars are not digits.")


    # Record the number of each type of device based on their role.
    nodes_roles_list = [ciq_db[node]["eNSESR Role"] for node in topo.nodes if node in ciq_db]
    nodes_cnt = {}
    for i in set([x for x in nodes_roles_list]):
        nodes_cnt[i] = len([x for x in nodes_roles_list if i in x])

    # Assign the node a correct position in the topology
    ebh_al_x = (((nodes_cnt["HUB AL"] * 2) + 2) / 2) - 2
    ebh_al_y = 6
    hub_bl_x = (((nodes_cnt["HUB AL"] * 2) + 2) / 2) - 2
    hub_bl_y = 5
    ptp_odd_x = (((nodes_cnt["HUB AL"] * 2) + 2) / 2) - 8
    ptp_odd_y = 4.75
    edn_odd_x = ptp_odd_x - 2
    edn_odd_y = ptp_odd_y + 0.25
    hub_bl_num = 0
    hub_al_num = 1
    hub_al_x = 1
    fixed_positions = {}
    # Assign the position of the hub nodes in the topology.
    for node in list(topo.nodes):
        if node in ciq_db:
            match ciq_db[node]["eNSESR Role"]:
                case "EBH AL":
                    fixed_positions.update({node: (ebh_al_x, ebh_al_y)})
                    ebh_al_x += 2
                    ebh_al_y += 0.25
                case "HUB BL":
                    fixed_positions.update({node: (hub_bl_x, hub_bl_y)})
                    hub_bl_num += 1
                    hub_bl_x += 2
                    hub_bl_y += 0.25
                case "HUB AL":
                    if hub_al_num == 1:
                        fixed_positions.update({node: (1, 2)})
                        hub_al_num += 1
                        hub_al_x += 2
                    elif hub_al_num == nodes_cnt["HUB AL"]:
                        fixed_positions.update({node: (1 + ((nodes_cnt["HUB AL"] - 1) * 2), 2.25)})
                    else:
                        if nodes_cnt["HUB AL"] % 2 == 1:
                            if hub_al_num < (nodes_cnt["HUB AL"] / 2):
                                fixed_positions.update({node: (hub_al_x, 2 - (0.25 * (hub_al_num - 1)))})
                            else:
                                fixed_positions.update({node: (hub_al_x, 2 - (0.25 * (nodes_cnt["HUB AL"] - hub_al_num)))})
                            hub_al_num += 1
                            hub_al_x += 2
                        elif nodes_cnt["HUB AL"] % 2 == 0:
                            if hub_al_num <= (nodes_cnt["HUB AL"] / 2):
                                fixed_positions.update({node: (hub_al_x, 2 - (0.25 * (hub_al_num - 1)))})
                            else:
                                fixed_positions.update({node: (hub_al_x, 2 - (0.25 * ((nodes_cnt["HUB AL"] - hub_al_num) - 1)))})
                            hub_al_num += 1
                            hub_al_x += 2
                case "DRAN SPOKE":
                    for intf in ciq_db[node]["Interfaces"]:
                        if ciq_db[node]["Interfaces"][intf]["description"].split('_')[0] in ciq_db:
                            hub_node = ciq_db[node]["Interfaces"][intf]["description"].split('_')[0]
                            if hub_node not in fixed_positions:
                                continue
                            x = fixed_positions[hub_node][0]
                            y = fixed_positions[hub_node][1] - 1
                            while (x, y) in fixed_positions.values():
                                    x += 0.5
                                    y -= 0.5
                            fixed_positions.update({node: (x, y)})

    ptp_even_x = hub_bl_x + 4
    ptp_even_y = hub_bl_y - (0.25 * (hub_bl_num + 1))
    edn_even_x = edn_odd_x - 2
    edn_even_y = edn_odd_y
    # Assign the position of the PTP nodes in the topology.
    for node in list(topo.nodes):
        if topo.nodes[node]["rank"] == node_ranking.index("PTP"):
            if int(node[-2:]) % 2 == 1:
                fixed_positions.update({node: (ptp_odd_x, ptp_odd_y)})
            elif int(node[-2:]) % 2 == 0:
                fixed_positions.update({node: (ptp_even_x, ptp_even_y)})
        elif topo.nodes[node]["rank"] == node_ranking.index("EDN"):
            if int(node[-2:]) % 2 == 1:
                fixed_positions.update({node: (edn_odd_x, edn_odd_y)})
            elif int(node[-2:]) % 2 == 0:
                fixed_positions.update({node: (edn_even_x, edn_even_y)})

    # This sets the max Axis range values so everything is in the topology output.
    x_min = 0
    y_min = 0
    x_max = 5
    y_max = 4
    for val in fixed_positions.values():
        if x_min > val[0] - 2:
            x_min = val[0] - 2
        if y_min > val[1] - 1:
            y_min = val[1] - 1
        if x_max < val[0] + 2:
            x_max = val[0] + 2
        if y_max < val[1] + 1:
            y_max = val[1] + 1

    matplotlib.use("Agg")                 # Force the topology creation to the backend.
    # Set the topology scale size.
    plt.figure(1, figsize=(x_max - x_min, y_max - y_min))
    # Set the X- and Y-Axis value ranges
    plt.xlim([x_min, x_max])
    plt.ylim([y_min, y_max])
    # Set the topology output to fixed using the data from above.
    pos = nx.spring_layout(topo, pos=fixed_positions, fixed=fixed_positions.keys())
    # Create the topology using the nodes and links from before.
    nx.draw_networkx(topo, pos=pos, with_labels=True, node_size=1000, font_size=8, label=site_name, clip_on=False, bbox=dict(boxstyle="square", facecolor="white", linewidth=0.3, alpha=1))
    nx.draw_networkx_edge_labels(topo, pos=pos, edge_labels=edge_labels)
    plt.title(site_name, fontsize=12)       # Add a title to the topology.

    # Display the topology.
    plt.savefig(os.path.join(topo_directory_name, f"{site_name} - Topology.png"))
    # plt.show()            # This command can be used to display the topology when troubleshooting.
    logger.info("Generating the Topology Diagram - Completed")


def create_folders(ciq_directory_name, site_name):
    # Capture the directory path where the NP Data will be stored.
    np_directory_name = os.path.join(ciq_directory_name, 'NP_DATA')
    # Check if the NP DATA directory has been created. If it hasn't, create it.
    if not os.path.isdir(os.path.join(ciq_directory_name, 'NP_DATA')):
        logger.info(f"Creating {np_directory_name}")
        os.mkdir(np_directory_name)

    # Capture the directory path where the site files will be stored.
    site_directory_name = os.path.join(ciq_directory_name, site_name + " - Files")
    # Check if the site directory has been created. If it hasn't, create it.
    if not os.path.isdir(site_directory_name):
        logger.info(f"Creating {site_directory_name}")
        os.mkdir(site_directory_name)

    # Capture the directory path where the configs will be stored.
    conf_directory_name = os.path.join(site_directory_name, "Configs")
    # Check if the configs directory has been created. If it hasn't, create it.
    if not os.path.isdir(conf_directory_name):
        logger.info(f"Creating {conf_directory_name}")
        os.mkdir(conf_directory_name)

    # Capture the directory path where the new pre- and post-check file will be stored.
    checks_directory_name = os.path.join(site_directory_name, "Checks")
    # Check if the check directory has been created. If it hasn't, create it.
    if not os.path.isdir(checks_directory_name):
        logger.info(f"Creating {checks_directory_name}")
        os.mkdir(checks_directory_name)

    # Capture the directory path where the new topology file will be stored.
    topo_directory_name = os.path.join(site_directory_name, "Topology")
    # Check if the topology directory has been created. If it hasn't, create it.
    if not os.path.isdir(topo_directory_name):
        logger.info(f"Creating {topo_directory_name}")
        os.mkdir(topo_directory_name)

    return np_directory_name, site_directory_name, conf_directory_name, checks_directory_name, topo_directory_name


def report_usage(cisco_device_count):
    logger.info("Reporting usage and savings to CX Catalog - Started")

    vzw_ebh_pid = "B89210"
    # Current Project ID for the Verizon Wireless EBH Account.
    pid_answer = input(f"The script has completed successfully.\n\nUsage and savings data is being logged to CX Catalog.\nWas this script used for PID {vzw_ebh_pid}? [yes] ")
    while pid_answer.lower() not in ("", "y", "yes", "n", "no"):
        pid_answer = input(f"Please enter yes or no.\nWas this script used for PID {vzw_ebh_pid}? [yes] ")
    if pid_answer.lower() in ["", "y", "yes"]:
        vzw_ebh_pid = "B89210"
    elif pid_answer.lower() in ["n", "no"]:
        vzw_ebh_pid = input("Provide the PID and press enter. ")

    savings_answer = input(f"{cisco_device_count} configs were generated.\nDid this save you {cisco_device_count*2} (2hrs per config) hours? [yes] ")
    while savings_answer.lower() not in ("", "y", "yes", "n", "no"):
        savings_answer = input(f"Please enter yes or no.\nDid this save you {cisco_device_count*2} (2hrs per config) hours? [yes] ")
    if savings_answer.lower() in ["", "y", "yes"]:
        saving_total = cisco_device_count * 2
    elif savings_answer.lower() in ["n", "no"]:
        saving_total = input("Provide the time savings, in hours, and press enter. ")

    try:
        # Submit telemetry data.
        telemetry_submission = aide.submit_statistics(
            pid=vzw_ebh_pid,  # This should be a valid PID
            tool_id="65d4dcb3e1617afe98fdec7f",
            metadata={
                "potential_savings": saving_total,  # Hours
                "report_savings": True,
            },
        )
    except:
        logger.warning("There was an issue reporting usage and savings to CX Catalog.")
    logger.info("Reporting usage and savings to CX Catalog - Completed")
    return


def main(ciq_file):
    # Extract the data from the CIQ.
    ciq_db, site_name = read_ciq(ciq_file)

    # Capture the directory path where the CIQ is stored and create all the required folders.
    np_directory_name, site_directory_name, conf_directory_name, checks_directory_name, topo_directory_name = create_folders(os.path.dirname(ciq_file), site_name)

    # Create a topology drawing based on the CIQ.
    build_topology(ciq_db, topo_directory_name, site_name)

    # Create a list of old hostnames from the CIQ.
    old_hosts_list = [ciq_db[x]['Old Hostname'] for x in ciq_db if 'Old Hostname' in ciq_db[x].keys() and '-cn-' not in ciq_db[x]["Old Hostname"].lower()]

    need_np_auth = False
    host_missing_config = []
    for host in old_hosts_list:
        if "n/a" == host.lower() or "na" == host.lower():
            continue
        else:
            # Create the file path used to save the old config.
            np_file_name = os.path.join(np_directory_name, f"{host}.cfg")
            # If the file isn't in the NP_DATA directory, add it to the host_missing_config list.
            if not os.path.isfile(np_file_name):
                host_missing_config.append(host)
                need_np_auth = True
                logger.warning(f"{host} config file not found locally. Adding this to the list of devices whose config will be pulled from NP.")

    if host_missing_config:
        today = datetime.today().date()
        last_week = today - timedelta(days=7)
        # Authenticate to NP once.
        auth = np_auth()
        # Retrieve the running config from NP for each device using their Old Hostname.
        for host in host_missing_config:
            # Create the file path used to save the old config.
            np_file_name = os.path.join(np_directory_name, f"{host}.cfg")
            logger.warning(f"Retrieving {host} config file from NP - Started")
            # Get config from NP.
            conf_contents = get_np(auth, host, today, last_week)
            # If the device isn't in NP, don't make a new file.
            if conf_contents == "":
                logger.critical(f"{host} config file not found in NP. Save the config file and then re-run the script.")
                logger.warning(f"Retrieving {host} config file from NP - Completed")
            elif conf_contents == "old":
                continue
            else:
                with open(np_file_name, 'w') as fd:
                    fd.write(conf_contents)
                logger.warning(f"Retrieving {host} config file from NP - Completed")
    else:
        # If the files already exist, inform the user that the local files will be used and not updated by NP.
        logger.info(f"Current config files already exists locally for each host. NP will not be used.")

    logger.info("The SNMP and Logging entries are based on ME-TRA-PR-14-0037 Cisco - All Franchise - Regional Management MOP v3.4.xlsx")
    mtso_config = {
        'APPLWIEE':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'CRDLIL13':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'CHNDINAA':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'CTTPMIBG':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'MSHWINBW':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'RYLOMICB':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'SFLDMILR':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'SFLEMIFX':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'AKROOH20':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'CNCQOH22':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'GAHGOHBT':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'LWCTOH02':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'MNTPOHAE':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'SCVIOHAG':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:5 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:16 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:A01:2130:A1:FEF:0:115 vrf management port default']},
        'BLTNMN86':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a05:2130:e0:fef:0:217 vrf management port default']},
        'GLVYMNNV':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a05:2130:e0:fef:0:217 vrf management port default']},
        'OMAJNEBM':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a05:2130:e0:fef:0:217 vrf management port default']},
        'OMALNEXU':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a05:2130:e0:fef:0:217 vrf management port default']},
        'OWTNMNCC':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a05:2130:e0:fef:0:217 vrf management port default']},
        'SXFLSDTU':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a05:2130:e0:fef:0:217 vrf management port default']},
        'MNRGKSAB':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a05:2130:e0:fef:0:217 vrf management port default']},
        'SPFDMOKC':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a05:2130:e0:fef:0:217 vrf management port default']},
        'SHPTLAWR':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:9 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:20 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:223 vrf management port default']},
        'CHRXNCLH':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'CLMASCMV':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'GNBQNC15':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'GNVLSCMZ':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'GRNRNCJB':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'NCHRSCPL':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'RLGHNCOR':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'ALPRGAGQ':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'BRHOALTB':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'DLTHGAGQ':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']},
        'MACNGAYQ':
            {'snmp': ['host 2001:4888:a02:2105:a0:fef:0:7 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210d:c0:fef:0:18 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a03:210a:c0:fef:0:203 traps version 2c 2Y2LHTZP31',
                      'host 2001:4888:a01:2106:a1:fef:0:203 traps version 2c 2Y2LHTZP31'],
             'logging': ['logging 2001:4888:a03:210c:c0:fef:0:187 vrf management port default']}
    }

    cisco_device_count = 0
    # Create the configuration file and check file for each host based on their new hostname.
    for host in ciq_db:
        # SKIP building Ciena configs.
        if "-cn-" in host.lower():
            continue
        # SKIP building Non-RAN Device configs.
        if "NON RAN" in ciq_db[host]["eNSESR Role"]:
            continue
        cisco_device_count += 1
        # Initiate the 'runfile' variable, so it is set as a list for future use.
        runfile = []
        if "n/a" == ciq_db[host]['Old Hostname'].lower() or "na" == ciq_db[host]['Old Hostname'].lower():
            logger.info(f"{ciq_db[host]['New Hostname']} - Device will be new, so we will not check for a config file in NP_DATA directory.")
            runfile = CiscoConfParse([""], factory=True)
        else:
            # Store the device's captured config to variable 'runfile' so it can be parsed as needed later on
            np_file_name = os.path.join(np_directory_name, f"{ciq_db[host]['Old Hostname']}.cfg")
            if os.path.isfile(np_file_name):
                runfile = CiscoConfParse(np_file_name, factory=True)
            else:
                logger.critical(f"{ciq_db[host]['New Hostname']} - Config file <{ciq_db[host]['Old Hostname']}.cfg> not found in NP_DATA directory. Save the raw running config to NP_DATA and re-run the script.")
                exit()
        # Get the current date in MMDDYYYY format.
        now = datetime.now()

        check_file_data = ""
        conf_file_data = ""

        match ciq_db[host]['eNSESR Role']:
            case "EBH AL":
                check_file_data = build_ebh_al().check(now, ciq_db[host])
                conf_file_data = build_ebh_al().config(ciq_db[host], ciq_db, runfile)
            case "HUB BL":
                check_file_data = build_hub_bl().check(now, ciq_db[host])
                conf_file_data = build_hub_bl().config(ciq_db[host], ciq_db, runfile, mtso_config)
            case "HUB AL":
                check_file_data = build_hub_al().check(now, ciq_db[host])
                conf_file_data = build_hub_al().config(ciq_db[host], ciq_db, runfile, mtso_config)
            case "DRAN SPOKE":
                check_file_data = build_dran_spoke().check(now, ciq_db[host])
                conf_file_data = build_dran_spoke().config(ciq_db[host], ciq_db, runfile, mtso_config)

        # Save the device's config change file in the site-specific directory.
        conf_file_name = os.path.join(conf_directory_name, f"{host}.cfg")
        logger.info(f"Generating {host} config file - Started")
        if not os.path.isfile(conf_file_name):
            # If the file doesn't already exist, create a new one.
            with open(conf_file_name, 'w', encoding="utf-8") as fd:
                fd.write(conf_file_data)
                logger.info(f"Generating {host} config file - Completed")
        else:
            # If the file already exist, inform the user that the local file will be updated.
            logger.info(f"Config file for {host} already exists locally. Contents will be updated.")
            with open(conf_file_name, 'w', encoding="utf-8") as fd:
                fd.write(conf_file_data)
                logger.info(f"Generating {host} config file - Completed")

        # Save the device's config change file in the site-specific directory.
        check_file_name = os.path.join(checks_directory_name, f"Checks - {host}.txt")
        logger.info(f"Generating {host} check file - Started")
        if not os.path.isfile(check_file_name):
            # If the file doesn't already exist, create a new one.
            with open(check_file_name, 'w', encoding="utf-8") as fd:
                fd.write(check_file_data)
                logger.info(f"Generating {host} check file - Completed")
        else:
            # If the file already exist, inform the user that the local file will be updated.
            logger.info(f"Check file for {host} already exists locally. Contents will be updated.")
            with open(check_file_name, 'w', encoding="utf-8") as fd:
                fd.write(check_file_data)
                logger.info(f"Generating {host} check file - Completed")

    # Save the ping validation commands to a file in the check directory.
    ping_data = hub_pings(ciq_db)
    # Save the hub's ping check file in the site-specific directory.
    check_file_name = os.path.join(checks_directory_name, f"Checks - Hub Pings.txt")
    logger.info(f"Generating Hub Pings check file - Started")
    if not os.path.isfile(check_file_name):
        # If the file doesn't already exist, create a new one.
        with open(check_file_name, 'w', encoding="utf-8") as fd:
            fd.write(ping_data)
            logger.info(f"Generating Hub Pings check file - Completed")
    else:
        # If the file already exist, inform the user that the local file will be updated.
        logger.info(f"Check file for Hub Pings already exists locally. Contents will be updated.")
        with open(check_file_name, 'w', encoding="utf-8") as fd:
            fd.write(ping_data)
            logger.info(f"Generating Hub Pings check file - Completed")

    # Add the contents of ciq_db to the log output.
    logger.info(f"{json.dumps(ciq_db, indent=2)}")

    report_usage(cisco_device_count)

    return


if __name__ == '__main__':
    # Log the time that the script started.
    logger.info("Script started.")
    logger.warning("WARNING: As of 9/7/2023, this script HAS been approved to use on iEN Hub migrations. The NCE is still responsible for reviewing all output files.")
    logger.warning("WARNING: The config files currently contain 'ssh server vrf CELL_MGMT' which is a workaround approved by Ray. Once the root issue is resolved, this will need to be removed from the automation.")
    logger.warning("WARNING: The NCE MUST review the configs and update the MTU for xNB and vDU interfaces to 1970. The automation doesn't currently do this.")
    logger.info("This script aligns with the following MOPs: 'EBH iEN Hub Migration to eNSE-SR MOP 1.5' & 'EBH eNSE SR Spoke Deployment GA MOP v1.2'.")

    current_directory_name = os.path.join(os.getcwd(), "")          # Capture the filepath to the current directory.
    files = os.listdir(current_directory_name)                      # Get list of files in the current directory.
    ciq_files = [x for x in files if "CIQ" in x and "xlsx" in x]    # Filter the files so only the CIQs are listed.
    ciq_files.sort(reverse=False)                                   # Sort the CIQ files alphabetically.

    # Provide a menu for the user to decide what CIQ to use.
    while True:
        # Find out how many CIQs are available and print the initial sentence.
        if len(ciq_files) == 1:
            print("The following CIQ file was found:")
        elif len(ciq_files) > 1:
            print("The following CIQ files were found:")
        else:
            print("No CIQ files were found. Please add a CIQ file to the directory and re-run the script.")
            exit()

        # Print out the CIQ options # and file name.
        for ciq in ciq_files:
            print(f"{ciq_files.index(ciq)} - {ciq}")

        # Record the users choice.
        ciq_choice = input(f"Which CIQ option do you want to use? [0-{len(ciq_files)-1}]")
        try:
            # This validates that the entry was an integer.
            ciq_choice = int(ciq_choice)
        except:
            # If the entry wasn't an integer, present the menu and prompt the user for their choice again.
            print("Invalid Option. Please try again.")
        else:
            if len(ciq_files) == 1:
                # If there is only 1 CIQ file, the only available option is '0'.
                if ciq_choice == 0:
                    break
                if ciq_choice != 0:
                    print("Invalid Option. Please try again.")
            else:
                # If there is more than 1 CIQ file, validate that the user entered an option within the available range.
                if ciq_choice in range(0, len(ciq_files), 1):
                    break
                else:
                    print("Invalid Option. Please try again.")

    # Capture the filepath to the CIQ file.
    ciq_file = os.path.join(current_directory_name, ciq_files[int(ciq_choice)])

    main(ciq_file=ciq_file)
    # Log the time that the script was completed.
    logger.info("Script ended.")
