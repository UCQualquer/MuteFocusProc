#include <Windows.h>
#include <tlhelp32.h>
#include "ProcessFamily.h";

HRESULT GetProcessFamily(DWORD dProcId, ProcessFamily** ppProcFamily)
{
	HRESULT hr = S_OK;
	HANDLE hTh32 = NULL;

	ProcessFamily* pProcFamily = new ProcessFamily{ 0 };
	
	pProcFamily->ProcessId = dProcId;
	pProcFamily->Children = new DWORD[0];

	PROCESSENTRY32 pe = { 0 };
	pe.dwSize = sizeof(PROCESSENTRY32);
	hTh32 = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

	auto maybeResize = [pProcFamily]()
		{
			// Expand array size by 8
			if (pProcFamily->ChildrenCount % 8 == 0)
			{
				DWORD* pNewArray = new DWORD[pProcFamily->ChildrenCount + 8]{ 0 };
				memcpy(pNewArray, pProcFamily->Children, pProcFamily->ChildrenCount);
				delete pProcFamily->Children;

				pProcFamily->Children = pNewArray;
			}
		};

	if (Process32First(hTh32, &pe)) {
		do {
			if (pe.th32ProcessID == dProcId) {
				pProcFamily->ParentId = pe.th32ParentProcessID;
			}
			else if (pe.th32ParentProcessID == dProcId)
			{
				maybeResize();
				pProcFamily->Children[pProcFamily->ChildrenCount++] = pe.th32ProcessID;
			}
		} while (Process32Next(hTh32, &pe));
	}
	else
		hr = E_FAIL;

	*ppProcFamily = pProcFamily;
	CloseHandle(hTh32);
	return hr;
}