# quickbooks-desktop-python
The QuickBooks SDK Python module provides a set of tools for integrating Python applications with QuickBooks Desktop. The module includes classes for connecting to the QuickBooks API, sending and receiving data using the QBXML protocol, and managing sessions with QuickBooks.


# Update
I no longer have quickbooks desktop but i need a copy of quickbooks desktop for this project.
I reached out to quickbooks and got a person to talk about it. They informed me, 
there is no way i could get an enterprise desktop installation to test the sdk and features. Secondly,
the trail version is very different, operation wise than the complete software. SDK integration will not work 
on the trail version. Please reach out to me if you can assist in this endeavor.

# Warning

This will install pandas and pywin32 automatically.
Secondly, this will assemble the QBSDK and install it locally.
Youve been warned.

# WorkFlow

1. Import libraries
2. Run precheck
   2.1 Check if Windows OS
   2.2 Check if installer for sdk is Ran
   2.3 Install necessary sdk components
3. Import Module



 # Example
 
```Python3
from QuickBooks import QuickBooks, XMLRequest

with QuickBooks() as qbxmlrp:
    response = qbxmlrp.process_request(XMLRequest.all_queries)
print(response)
```

# The Goal
 Make it easy for developers to access and manipulate QuickBooks data, 
 while providing robust error handling, management,
 and support for working with XML data. Our goal is to help developers save time and
 resources when working with QuickBooks data, and to enable them to build custom 
 integrations that meet their unique business needs. 
 
 # Why Python?
First, Python is a popular and powerful language that is widely used in the software development community. It has a large and active user base.

Second, Python is  a great choice for building custom integrations and working with data from external sources like QuickBooks. Python also has a wide range of libraries and modules that can be easily integrated into a project, which can save developers a significant amount of time and effort.

Third, Python has a clean and simple syntax, which makes it easy to read and write code. This can help developers write code more quickly and efficiently, and can also make it easier for other developers to understand and modify the code as needed.

# Getting Started
```PowerShell
git clone https://github.com/Donny-GUI/quickbooks-desktop-python.git
cd quickbooks-desktop-python
python qbdesktop.py
```
