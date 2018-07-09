# Get-ExchangePQT

The script is an alternative to the Processor Query Tool and automates the steps 
to calculate the SPECint 2006 Rate Value (www.spec.org) of your planned
processor model when Exchange Server 2010/2013/2016 configurations. This value
can be used for the calculator, or you can use it to find a system which conforms
to a specific SPECint throughput.

For virtualized environments, you can determine the SPECint for a specific 
virtual processor ratio in combination with allocated number of vCPUs.

### Prerequisites

None

### Usage

Search all specs for systems containing x3430 CPUs
```
.\Get-ExchangePQT.ps1 -CPU x3430 
```

### About

-

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

 