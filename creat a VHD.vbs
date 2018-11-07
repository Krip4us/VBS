On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
errReturn = objVS.CreateDynamicVirtualHardDisk _
    ("C:\Virtual Machines\Disks\Scripted_HardDisk.vhd", 20)
