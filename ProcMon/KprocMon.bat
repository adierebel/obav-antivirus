;@echo off
;goto make

;===================================================================================================

.386
.model flat, stdcall
option casemap:none

;===================================================================================================

include c:\masm32\include\w2k\ntstatus.inc
include c:\masm32\include\w2k\ntddk.inc
include c:\masm32\include\w2k\ntoskrnl.inc
include c:\masm32\Macros\Strings.mac
include common.inc

includelib c:\masm32\lib\w2k\ntoskrnl.lib

;===================================================================================================

.const

CCOUNTED_UNICODE_STRING	"\\Device\\KprocMon", g_usDeviceName, 4
CCOUNTED_UNICODE_STRING	"\\DosDevices\\KprocMon", g_usSymbolicLinkName, 4

.data?

g_pkEventObject		PKEVENT		?
g_dwImageFileNameOffset	DWORD		?
g_fbNotifyRoutineSet	BOOL		?

g_ProcessData		PROCESS_DATA	<>


.code

;===================================================================================================

DispatchCreateClose proc pDeviceObject:PDEVICE_OBJECT, pIrp:PIRP

	mov ecx, pIrp
	mov (_IRP PTR [ecx]).IoStatus.Status, STATUS_SUCCESS
	and (_IRP PTR [ecx]).IoStatus.Information, 0

	fastcall IofCompleteRequest, ecx, IO_NO_INCREMENT

	mov eax, STATUS_SUCCESS
	ret

DispatchCreateClose endp

;===================================================================================================

ProcessNotifyRoutine proc dwParentId:DWORD, dwProcessId:DWORD, bCreate:BOOL

local peProcess:PVOID
local fbDereference:BOOL
local us:UNICODE_STRING
local as:ANSI_STRING

	push eax
	invoke PsLookupProcessByProcessId, dwProcessId, esp
	pop peProcess
	.if eax == STATUS_SUCCESS
		mov fbDereference, TRUE
	.else
		invoke IoGetCurrentProcess
		mov peProcess, eax
		and fbDereference, FALSE
	.endif

	mov eax, dwProcessId
	mov g_ProcessData.dwProcessId, eax

	mov eax, bCreate
	mov g_ProcessData.bCreate, eax

	.if fbDereference
		fastcall ObfDereferenceObject, peProcess
	.endif

	invoke KeSetEvent, g_pkEventObject, 0, FALSE
	ret

ProcessNotifyRoutine endp

;===================================================================================================

DispatchControl proc uses esi edi pDeviceObject:PDEVICE_OBJECT, pIrp:PIRP
local liDelayTime:LARGE_INTEGER

	mov esi, pIrp
	assume esi:ptr _IRP

	mov [esi].IoStatus.Status, STATUS_UNSUCCESSFUL
	and [esi].IoStatus.Information, 0

	IoGetCurrentIrpStackLocation esi
	mov edi, eax
	assume edi:ptr IO_STACK_LOCATION

	.if [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_SET_NOTIFY
		.if [edi].Parameters.DeviceIoControl.InputBufferLength >= sizeof HANDLE
			.if g_fbNotifyRoutineSet == FALSE

				mov edx, [esi].AssociatedIrp.SystemBuffer
				mov edx, [edx]			; user-mode hEvent

				mov ecx, ExEventObjectType
				mov ecx, [ecx]
				mov ecx, [ecx]			; PTR OBJECT_TYPE

				invoke ObReferenceObjectByHandle, edx, EVENT_MODIFY_STATE, ecx, UserMode, addr g_pkEventObject, NULL
				.if eax == STATUS_SUCCESS
					invoke PsSetCreateProcessNotifyRoutine, ProcessNotifyRoutine, FALSE
					mov [esi].IoStatus.Status, eax
					.if eax == STATUS_SUCCESS
						mov g_fbNotifyRoutineSet, TRUE

						mov eax, pDeviceObject
						mov eax, (DEVICE_OBJECT PTR [eax]).DriverObject
						and (DRIVER_OBJECT PTR [eax]).DriverUnload, NULL
					.endif
				.else
					mov [esi].IoStatus.Status, eax
				.endif
			.endif
		.else
			mov [esi].IoStatus.Status, STATUS_BUFFER_TOO_SMALL
		.endif

	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_REMOVE_NOTIFY
		.if g_fbNotifyRoutineSet == TRUE
			invoke PsSetCreateProcessNotifyRoutine, ProcessNotifyRoutine, TRUE
			mov [esi].IoStatus.Status, eax
			.if eax == STATUS_SUCCESS
				and g_fbNotifyRoutineSet, FALSE

				or liDelayTime.HighPart, -1
				mov liDelayTime.LowPart, -1000000
				invoke KeDelayExecutionThread, KernelMode, FALSE, addr liDelayTime

				; Make driver unloadable
				mov eax, pDeviceObject
				mov eax, (DEVICE_OBJECT PTR [eax]).DriverObject
				mov (DRIVER_OBJECT PTR [eax]).DriverUnload, offset DriverUnload

				.if g_pkEventObject != NULL
					invoke ObDereferenceObject, g_pkEventObject
					and g_pkEventObject, NULL
				.endif
			.endif
		.endif

	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_GET_PROCESS_DATA
		.if [edi].Parameters.DeviceIoControl.OutputBufferLength >= sizeof PROCESS_DATA

			mov eax, [esi].AssociatedIrp.SystemBuffer
			invoke memcpy, eax, offset g_ProcessData, sizeof g_ProcessData

			mov [esi].IoStatus.Status, STATUS_SUCCESS
			mov [esi].IoStatus.Information, sizeof g_ProcessData
		.else
			mov [esi].IoStatus.Status, STATUS_BUFFER_TOO_SMALL
		.endif
	.else
		mov [esi].IoStatus.Status, STATUS_INVALID_DEVICE_REQUEST
	.endif

	push [esi].IoStatus.Status
	assume edi:nothing
	assume esi:nothing
	fastcall IofCompleteRequest, esi, IO_NO_INCREMENT

	pop eax			; [esi].IoStatus.Status
	ret

DispatchControl endp


;===================================================================================================

DriverUnload proc pDriverObject:PDRIVER_OBJECT

	invoke IoDeleteSymbolicLink, addr g_usSymbolicLinkName

	mov eax, pDriverObject
	invoke IoDeleteDevice, (DRIVER_OBJECT PTR [eax]).DeviceObject

	ret

DriverUnload endp

;===================================================================================================

.code INIT

;===================================================================================================

DriverEntry proc pDriverObject:PDRIVER_OBJECT, pusRegistryPath:PUNICODE_STRING

local status:NTSTATUS
local pDeviceObject:PDEVICE_OBJECT

	mov status, STATUS_DEVICE_CONFIGURATION_ERROR

	invoke IoCreateDevice, pDriverObject, 0, addr g_usDeviceName, \
				FILE_DEVICE_UNKNOWN, 0, TRUE, addr pDeviceObject
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

			and g_fbNotifyRoutineSet, FALSE
			invoke memset, addr g_ProcessData, 0, sizeof g_ProcessData
		
			mov status, STATUS_SUCCESS
		.else
			invoke IoDeleteDevice, pDeviceObject
		.endif
	.endif

	mov eax, status
	ret

DriverEntry endp

;===================================================================================================

end DriverEntry

;===================================================================================================

:make

set drv=KprocMon

	c:\masm32\bin\rc /v rsrc.rc
	c:\masm32\bin\cvtres /machine:ix86 rsrc.res
	if errorlevel 0 goto final
	pause
	exit
:final
if exist rsrc.res del rsrc.res
c:\masm32\bin\ml /nologo /c /coff %drv%.bat
c:\masm32\bin\link /nologo /driver /base:0x10000 /align:32 /out:%drv%.sys /subsystem:native /ignore:4078 %drv%.obj  rsrc.obj

del %drv%.obj
del rsrc.obj

echo.
pause

;===================================================================================================