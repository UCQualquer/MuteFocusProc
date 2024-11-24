#pragma once

#include <Windows.h>

struct ProcessFamily {
	DWORD ProcessId;
	DWORD ParentId;
	UINT ChildrenCount;
	DWORD* Children;
};

HRESULT GetProcessFamily(DWORD dProcId, ProcessFamily** pProcFamily);