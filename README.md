# getm365endpoints

## About
Microsoft publishes a web page of all M365 endpoints which are the set of destination IP addresses, ports, DNS domain names, and URLs required by users to access M365 cloud-based services on the Internet.  Microsoft makes available a REST API Web Service to help customers programmatically consume this information.

Endpoints:   https://aka.ms/o365endpoints

Web service: https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service

The default output from the web service however may not be in a format you require, or easy to consume.  For example, the IPs and URLs are listed by ID and the same network range can be shown several times in the list.

This script was developed to simplify and automate the process of checking for updates, retrieving the latest endpoints, and sorting and filtering the output into various easy to read outputs and formats.

It creates several files including:
- an Excel spreadsheet detailing IPs and URLs by category (Optimize, Allow, Default), by Required or Optional, and by port (requires Excel)
- text files of IPs and URLs by category, service area, port, etc.
- text file of URLs in .pac file format
- text files of IPs by CIDR notation, subnet mask notation, or wildcard mask notation
- csv files of all current IP and URLs, change logs between versions, and notes, etc.

## Snippet of spreadsheet created:
![image](https://user-images.githubusercontent.com/99920431/215800088-b327db21-b4cd-470a-8697-ee0a2cdb70f2.png)

## Updates to IPs and URLs
Updates to the M365 endpoint lists are made approximately monthly at the end of the month, so subscribing to the RSS feed will ensure you receive emails from Microsoft when any changes to the endpoints are made.  This way you only need to run the script when you get an alert that the endpoint list has been updated.
You may also want to create a rule or flow to be alerted when the RSS feed is updated.

## Microsoft Guidance on endpoints and Optimization methods:
Optimize endpoints are required for connectivity to every Office 365 service and represent over 75% of Office 365 bandwidth, connections, and volume of data
These endpoints are the most sensitive to network performance, latency, and availability

OPTIMIZE:
Bypass network devices and services that perform traffic interception, SSL decryption, deep packet inspection, and content filtering
Bypass on-premises proxy devices and cloud-based proxy services
Facilitate direct connectivity for VPN users by implementing split tunneling

ALLOW:
Bypass endpoints on network devices and services that perform traffic interception, SSL decryption, deep packet inspection, and content filtering

DEFAULT:
No optimization required - treat like any other internet traffic

## Notes
Note: This list of published IPs/URLs from Microsoft represents requirements for connectivity from a user's machine to Office 365.  There are additional lists of IPs/URLs for connectivity that may also be required for additional services, features and applications if being used, which are noted in the following pages:

Additional endpoints not included in the Office 365 IP Address and URL Web service
https://learn.microsoft.com/en-us/microsoft-365/enterprise/additional-office365-ip-addresses-and-urls?view=o365-worldwide

Hybrid Identity required ports and protocols
https://learn.microsoft.com/en-us/azure/active-directory/hybrid/reference-connect-ports

Azure Logic Apps
https://learn.microsoft.com/en-us/azure/logic-apps/logic-apps-limits-and-config?tabs=consumption%2Cazure-portal

Teams Direct Routing
https://learn.microsoft.com/en-us/microsoftteams/direct-routing-plan#sip-signaling-fqdns

Microsoft Intune
https://learn.microsoft.com/en-us/mem/intune/fundamentals/intune-endpoints

Microsoft Defender for Cloud Apps (previously MCAS)
https://learn.microsoft.com/en-us/defender-cloud-apps/network-requirements

Network requests in Office for Mac
https://learn.microsoft.com/en-us/microsoft-365/enterprise/network-requests-in-office-2016-for-mac?view=o365-worldwide

Troubleshooting live events in Microsoft Stream (Classic)
https://learn.microsoft.com/en-us/stream/live-event-retirement#make-sure-all-endpoints-are-reachable-in-your-network

How to Optimize Stream & Live Events traffic in a VPN scenario
https://techcommunity.microsoft.com/t5/microsoft-365-blog/how-to-optimize-stream-amp-live-events-traffic-in-a-vpn-scenario/ba-p/1439767

Yammer
https://techcommunity.microsoft.com/t5/yammer-blog/using-hard-coded-ip-addresses-for-yammer-is-not-recommended/ba-p/276592

Windows Notification Service
https://learn.microsoft.com/en-us/windows/apps/design/shell/tiles-and-notifications/firewall-allowlist-config

WSUS
https://learn.microsoft.com/en-us/windows-server/administration/windows-server-update-services/deploy/2-configure-wsus#211-connection-from-the-wsus-server-to-the-internet
