#include <ntddk.h>

NTSTATUS AddDevice(PDRIVER_OBJECT DriverObject,PDEVICE_OBJECT pdo)
{
	UNICODE_STRING name;
	PDEVICE_OBJECT fdo;
	NTSTATUS status;
	RtlInitUnicodeString(&name,L"\\Device\\HelloWDM");
	status =IoCreateDevice(DriverObject,
								 0,
								 &name,
								 FILE_DEVICE_UNKNOWN,
								 0,
								 TRUE,
								 &fdo);
	if(!NT_SUCCESS(status))
	{
		DbgPrint("Create Device failure\n");
		return status;
	}
	return STATUS_SUCCESS;
}

NTSTATUS HelloWDM(PDEVICE_OBJECT pdo,PIRP irp)
{
	return STATUS_SUCCESS;
}

NTSTATUS DriverEntry(IN PDRIVER_OBJECT DriverObject,IN PUNICODE_STRING RegistryPath)
{
	DriverObject->DriverExtension->AddDevice           = AddDevice;
	DriverObject->MajorFunction[IRP_MJ_SYSTEM_CONTROL] = HelloWDM;
	return STATUS_SUCCESS;
}