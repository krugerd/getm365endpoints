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
