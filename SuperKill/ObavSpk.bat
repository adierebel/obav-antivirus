;@echo off
;goto make
; ---------------------------------oj4nBL4NK--------------------------------------------------------
.386
.model flat, stdcall
option casemap:none

include c:\masm32\include\w2k\ntstatus.inc
include c:\masm32\include\w2k\ntddk.inc
include c:\masm32\include\w2k\ntoskrnl.inc
include common.inc
include c:\masm32\Macros\Strings.mac

includelib c:\masm32\lib\w2k\ntoskrnl.lib

.const
CCOUNTED_UNICODE_STRING	"\\Device\\ojansuperkill", g_usDeviceName, 4
CCOUNTED_UNICODE_STRING	"\\DosDevices\\ojansuperkill", g_usSymbolicLinkName, 4

; ---------------------------------oj4nBL4NK--------------------------------------------------------
.code

; ---------------------------------oj4nBL4NK--------------------------------------------------------
DispatchCreateClose proc pDeviceObject:PDEVICE_OBJECT, pIrp:PIRP
	mov eax, pIrp
	assume eax:ptr _IRP
	mov [eax].IoStatus.Status, STATUS_SUCCESS
	and [eax].IoStatus.Information, 0
	assume eax:nothing

	fastcall IofCompleteRequest, pIrp, IO_NO_INCREMENT

	mov eax, STATUS_SUCCESS
	ret
DispatchCreateClose endp

; ---------------------------------oj4nBL4NK--------------------------------------------------------

DispatchControl proc uses esi edi pDeviceObject:PDEVICE_OBJECT, pIrp:PIRP

local status:NTSTATUS
local dwBytesReturned:DWORD
local ObjectAttributes:OBJECT_ATTRIBUTES
local ProcessHandle   :DWORD
local ClientId        :CLIENT_ID

	and dwBytesReturned, 0

	mov esi, pIrp
	assume esi:ptr _IRP

	IoGetCurrentIrpStackLocation esi
	mov edi, eax
	assume edi:ptr IO_STACK_LOCATION

	.if [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_KILL_PROCCESS
			mov edi, [esi].AssociatedIrp.SystemBuffer
			assume edi:ptr DWORD

			mov ClientId.UniqueThread, 0
			mov eax,[edi][0*(sizeof DWORD)]
			mov ClientId.UniqueProcess, eax
			InitializeObjectAttributes addr ObjectAttributes, NULL, OBJ_KERNEL_HANDLE + OBJ_CASE_INSENSITIVE, 0, NULL
			invoke ZwOpenProcess,addr ProcessHandle, PROCESS_ALL_ACCESS, addr ObjectAttributes, addr ClientId
			invoke ZwTerminateProcess,ProcessHandle, 0
			invoke ZwClose,ProcessHandle

			mov status, STATUS_SUCCESS
	.else
		mov status, STATUS_INVALID_DEVICE_REQUEST
	.endif
	assume edi:nothing

	push status
	pop [esi].IoStatus.Status

	push dwBytesReturned
	pop [esi].IoStatus.Information

	assume esi:nothing

	fastcall IofCompleteRequest, esi, IO_NO_INCREMENT

	mov eax, status
	ret

DispatchControl endp

; ---------------------------------oj4nBL4NK--------------------------------------------------------

DriverUnload proc pDriverObject:PDRIVER_OBJECT
	invoke IoDeleteSymbolicLink, addr g_usSymbolicLinkName
	mov eax, pDriverObject
	invoke IoDeleteDevice, (DRIVER_OBJECT PTR [eax]).DeviceObject
	ret
DriverUnload endp

; ---------------------------------oj4nBL4NK--------------------------------------------------------

.code INIT

; ---------------------------------oj4nBL4NK--------------------------------------------------------

DriverEntry proc pDriverObject:PDRIVER_OBJECT, pusRegistryPath:PUNICODE_STRING
local status:NTSTATUS
local pDeviceObject:PDEVICE_OBJECT

	mov status, STATUS_DEVICE_CONFIGURATION_ERROR

	invoke IoCreateDevice, pDriverObject, 0, addr g_usDeviceName, FILE_DEVICE_UNKNOWN, 0, FALSE, addr pDeviceObject
	.if eax == STATUS_SUCCESS
		invoke IoCreateSymbolicLink, addr g_usSymbolicLinkName, addr g_usDeviceName
		.if eax == STATUS_SUCCESS
			mov eax, pDriverObject
			assume eax:ptr DRIVER_OBJECT
			mov [eax].MajorFunction[IRP_MJ_CREATE*(sizeof PVOID)],			offset DispatchCreateClose
			mov [eax].MajorFunction[IRP_MJ_CLOSE*(sizeof PVOID)],			offset DispatchCreateClose
			mov [eax].MajorFunction[IRP_MJ_DEVICE_CONTROL*(sizeof PVOID)],	offset DispatchControl
			mov [eax].DriverUnload,											offset DriverUnload
			assume eax:nothing
			mov status, STATUS_SUCCESS
		.else
			invoke IoDeleteDevice, pDeviceObject
		.endif
	.endif

	mov eax, status
	ret

DriverEntry endp

; ---------------------------------oj4nBL4NK--------------------------------------------------------

end DriverEntry

:make

set drv=ObavSpk

:makerc
if exist rsrc.obj goto final
	c:\masm32\bin\rc /v rsrc.rc
	c:\masm32\bin\cvtres /machine:ix86 rsrc.res
	if errorlevel 0 goto final
		pause
		exit
:final
if exist rsrc.res del rsrc.res
c:\masm32\bin\ml /nologo /c /coff %drv%.bat
c:\masm32\bin\link /nologo /driver /base:0x10000 /align:32 /out:%drv%.sys /subsystem:native /ignore:4078 %drv%.obj rsrc.obj

del %drv%.obj

pause
