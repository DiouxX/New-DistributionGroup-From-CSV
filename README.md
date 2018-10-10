# Table of Contents
- [Table of Contents](#table-of-contents)
- [Description](#description)
- [Usage](#usage)
    - [Mandatory parameter](#mandatory-parameter)
    - [Optionnal Parameter](#optionnal-parameter)
- [Script excecution](#script-excecution)
    - [Basic Execution](#basic-execution)
    - [Specified Delimiter](#specified-delimiter)
    - [Specified OU](#specified-ou)
- [Praticla CSV](#praticla-csv)

# Description
Script to create Distribution Group on Exchaneg Server from CSV file
If Distribution Group don't exist, then created

# Usage
## Mandatory parameter

-PathCsv : Specified CSV file
-HeaderDistributionGroup : Field to identify Distribution Group name 
-HeaderMailAddress : Field to identify mail address to add on Distribution Group

## Optionnal Parameter

-OrganizationalUnit : Field to specified Active Directory OU to created Distribution Group
-Delimiter : By default, CSV delimiter is ';' character but if you want, you can specified another character

# Script excecution
## Basic Execution

```powershell
.\New-DistributionGroup-From-CSV.ps1 -PathCsv .\your-file.csv -HeaderDistributionGroup DistributionGroup -HeaderMailAddress Mail
```

## Specified Delimiter

```powershell
.\New-DistributionGroup-From-CSV.ps1 -PathCsv .\your-file.csv -HeaderDistributionGroup DistributionGroup -HeaderMailAddress Mail -Delimiter ","
```

## Specified OU

```powershell
.\New-DistributionGroup-From-CSV.ps1 -PathCsv .\your-file.csv -HeaderDistributionGroup DistributionGroup -HeaderMailAddress Mail -OrganizationalUnit "Distribution Group" -Delimiter ","
```

# Practical CSV
```
Name;E-Mail;DistributionGroup
User1 Name;user1.name@domain.com;Distribution_group1
User2 Name;user1.name@domain.com;Distribution_group1
User3 Name;user1.name@domain.com;Distribution_group2
User4 Name;user1.name@domain.com;Distribution_group1
User5 Name;user1.name@domain.com;Distribution_group2
User6 Name;user1.name@domain.com;Distribution_group3
```

With this example, use the following command :

```powershell
.\New-DistributionGroup-From-CSV.ps1 -PathCsv .\your-file.csv -HeaderDistributionGroup DistributionGroup -HeaderMailAddress E-Mail
```
