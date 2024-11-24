// MuteFocusProc.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#ifdef _DEBUG
#include <iostream>
#define LOG(...) fprintf(stdout, __VA_ARGS__)
#define LOG_ERR(...) fprintf(stderr, __VA_ARGS__)
#else
#define LOG(...)
#define LOG_ERR(...)
#endif

#define EXIT_ON_ERROR(hr, ln) if (FAILED(hr)) { LOG_ERR("Error on line %d\n", ln); goto Exit; }
#define EXIT_ON_NULL(hndl, ln) if (hndl == NULL) { LOG_ERR("Error on line %d\n", ln); goto Exit; }
#define SAFE_RELEASE(punk) if ((punk) != NULL) { (punk)->Release(); (punk) = NULL; }

#include <Windows.h>
#include <Oleacc.h>
#include <audioclient.h>
#include <audiopolicy.h>
#include <endpointvolume.h>
#include <mmdeviceapi.h>
#include <Functiondiscoverykeys_devpkey.h>
#include <tlhelp32.h>
#include "ProcessFamily.h";

// https://github.com/microsoft/Windows-classic-samples/blob/d338bb385b1ac47073e3540dbfa810f4dcb12ed8/Samples/Win7Samples/multimedia/mediafoundation/MFPlayer2/AudioSessionVolume.cpp#L30
// {2715279F-4139-4ba0-9CB1-B351F1B58A4A}
static const GUID AudioSessionVolumeCtx = { 0x2715279f, 0x4139, 0x4ba0, { 0x9c, 0xb1, 0xb3, 0x51, 0xf1, 0xb5, 0x8a, 0x4a } };

HRESULT GetFocusedProcessId(DWORD* dpProcId)
{
	HWND wFocusedWind;

	wFocusedWind = GetForegroundWindow();
	if (wFocusedWind == NULL) {
		LOG_ERR("GetForegroundWindow fail\n");
		return E_FAIL;
	}

	GetWindowThreadProcessId(wFocusedWind, dpProcId);
	if (dpProcId == NULL) {
		LOG_ERR("GetWindowThreadProcessId fail\n");
		return E_FAIL;
	}
	
	return S_OK;
}

DWORD GetParentProcessId(DWORD dProcId)
{
	HANDLE h = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	PROCESSENTRY32 pe = { 0 };
	pe.dwSize = sizeof(PROCESSENTRY32);

	if (Process32First(h, &pe)) {
		do {
			if (pe.th32ProcessID == dProcId) {
				CloseHandle(h);
				return pe.th32ParentProcessID;
			}
		} while (Process32Next(h, &pe));
	}

	CloseHandle(h);
	return 0;
}

BOOL IsProcessRelated(DWORD dProcId, DWORD dProcId1)
{
	HANDLE h = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	PROCESSENTRY32 pe = { 0 };
	pe.dwSize = sizeof(PROCESSENTRY32);

	if (Process32First(h, &pe)) {
		do {
			if (pe.th32ParentProcessID == dProcId && pe.th32ProcessID == dProcId1
				|| pe.th32ParentProcessID == dProcId1 && pe.th32ProcessID == dProcId) {
				CloseHandle(h);
				return TRUE;
			}
		} while (Process32Next(h, &pe));
	}

	CloseHandle(h);
	return FALSE;
}

void LogDeviceEndpoint(IMMDevice* pEndpoint)
{
	HRESULT hr;

	LPWSTR pwszId;
	IPropertyStore* pProps = NULL;
	PROPVARIANT varName;

	// Print endpoint id
	hr = pEndpoint->GetId(&pwszId);
	if (SUCCEEDED(hr))
		LOG("Device Id: '%S'\n", pwszId);

	// Print endpoint friendly name
	hr = pEndpoint->OpenPropertyStore(STGM_READ, &pProps);
	if (SUCCEEDED(hr))
	{
		PropVariantInit(&varName);

		hr = pProps->GetValue(PKEY_Device_FriendlyName, &varName);
		if (SUCCEEDED(hr) && varName.vt != VT_EMPTY)
			LOG("Friendly Name: '%S'\n", varName.pwszVal);
	}

	CoTaskMemFree(pwszId);
	PropVariantClear(&varName);
	SAFE_RELEASE(pProps);
}

void LogAudioSession(IAudioSessionControl2* pSessionControl2)
{
	HRESULT hr;

	LPWSTR dDispName;
	LPWSTR dIconPath;
	DWORD dProcId;

	hr = pSessionControl2->GetProcessId(&dProcId);
	if (SUCCEEDED(hr))
		LOG("- Session Process Id: %d\n", dProcId);

	hr = pSessionControl2->GetDisplayName(&dDispName);
	if (SUCCEEDED(hr) && wcslen(dDispName) > 0)
		LOG("- Display Name: '%S'\n", dDispName);

	hr = pSessionControl2->GetIconPath(&dIconPath);
	if (SUCCEEDED(hr) && wcslen(dIconPath) > 0)
		LOG("- Icon Path: '%S'\n", dIconPath);

	CoTaskMemFree(dDispName);
	CoTaskMemFree(dIconPath);
}

HRESULT GetProcessName(DWORD dProcId, LPWSTR lpProcName)
{
	HRESULT hr = E_FAIL;
	HANDLE handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, dProcId);

	if (handle)
	{
		DWORD buffSize = 1024;
		if (QueryFullProcessImageNameW(handle, 0, lpProcName, &buffSize))
		{
			hr = S_OK;
		}
	}

	CloseHandle(handle);
	return hr;
}

HRESULT FindPreferredAudioSession(DWORD dProcId, IMMDevice* pEndpoint, IAudioSessionControl** pBestSessionControl)
{
	HRESULT hr;
	LPWSTR lpwProcName = new WCHAR[1024];
	ProcessFamily* pProcFamily = NULL;

	IAudioSessionManager2* pAudioSessionManager = NULL;
	IAudioSessionEnumerator* pSessionEnumerator = NULL;

	IAudioSessionControl* pSessionControl = NULL;
	IAudioSessionControl2* pSessionControl2 = NULL;

	IAudioSessionControl* pFoundBestSessionControl = NULL;

	IAudioSessionControl** pSessions = NULL;
	IAudioSessionControl2** pSessions2 = NULL;

	int sessionCount = 0;

	hr = GetProcessName(dProcId, lpwProcName);
	EXIT_ON_ERROR(hr, __LINE__);

	LOG("Target process executable name: '%S'\n", lpwProcName);

	// Get the IAudioSessionManager2 interface
	hr = pEndpoint->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, NULL, (void**)&pAudioSessionManager);
	EXIT_ON_ERROR(hr, __LINE__);

	// Enumerate audio sessions
	hr = pAudioSessionManager->GetSessionEnumerator(&pSessionEnumerator);
	EXIT_ON_ERROR(hr, __LINE__);

	hr = pSessionEnumerator->GetCount(&sessionCount);
	EXIT_ON_ERROR(hr, __LINE__);

	// Get list of audio sessions. Better to do it once than to enumerate
	// for every check we do.
	pSessions = (IAudioSessionControl**)malloc(sizeof(IAudioSessionControl*) * sessionCount);
	pSessions2 = (IAudioSessionControl2**)malloc(sizeof(IAudioSessionControl2*) * sessionCount);
	for (int sessionIndex = 0; sessionIndex < sessionCount; sessionIndex++) {
		// Get controller v1
		hr = pSessionEnumerator->GetSession(sessionIndex, &pSessionControl);
		if (FAILED(hr)) continue;

		// Get controller v2
		hr = pSessionControl->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&pSessionControl2);
		if (FAILED(hr)) continue;

		pSessions[sessionIndex] = pSessionControl;
		pSessions2[sessionIndex] = pSessionControl2;
		SAFE_RELEASE(pSessionControl);
	}

	// Try finding audio session with PID
	for (int sessionIndex = 0; sessionIndex < sessionCount && pFoundBestSessionControl == NULL; sessionIndex++)
	{
		pSessionControl = pSessions[sessionIndex];
		pSessionControl2 = pSessions2[sessionIndex];

		// Get audio session process id
		DWORD dSessionProcId;
		hr = pSessionControl2->GetProcessId(&dSessionProcId);
		if (FAILED(hr) || dSessionProcId == 0) continue;

		// Match process id
		if (dSessionProcId == dProcId)
		{
			LOG("Found session by matching PID: %d\n", dSessionProcId);
			pFoundBestSessionControl = pSessionControl;
			break;
		}
	}

	// Try with executable name
	for (int sessionIndex = 0; sessionIndex < sessionCount && pFoundBestSessionControl == NULL; sessionIndex++)
	{
		pSessionControl = pSessions[sessionIndex];
		pSessionControl2 = pSessions2[sessionIndex];

		// Get audio session process id
		DWORD dSessionProcId;
		hr = pSessionControl2->GetProcessId(&dSessionProcId);
		if (FAILED(hr) || dSessionProcId == 0) continue;

		// Match process executable
		LPWSTR lpwSessProcName = new WCHAR[1024]{ 0 };
		hr = GetProcessName(dSessionProcId, lpwSessProcName);
		if (SUCCEEDED(hr)) {
			if (wcscmp(lpwProcName, lpwSessProcName) == 0)
			{
				LOG("Found session by matching executable name: '%S'\n", lpwSessProcName);
				pFoundBestSessionControl = pSessionControl;
				break;
			}
		}
	}

	// Try with child processes
	for (int sessionIndex = 0; sessionIndex < sessionCount && pFoundBestSessionControl == NULL; sessionIndex++)
	{
		pSessionControl = pSessions[sessionIndex];
		pSessionControl2 = pSessions2[sessionIndex];

		if (pProcFamily == NULL) // Initialize only once
			hr = GetProcessFamily(dProcId, &pProcFamily);

		// Get audio session process id
		DWORD dSessionProcId;
		hr = pSessionControl2->GetProcessId(&dSessionProcId);
		if (FAILED(hr) || dSessionProcId == 0) continue;

		// Search for matching child process
		for (int i = 0; i < pProcFamily->ChildrenCount && pFoundBestSessionControl == NULL; i++)
		{
			DWORD dChildId = pProcFamily->Children[i];
			if (dChildId == dSessionProcId) // The session owner process is a child process of our target
			{
				LOG("Found session by matching child process: %d\n", dChildId);
				pFoundBestSessionControl = pSessionControl;
			}
		}

		if (pFoundBestSessionControl != NULL)
			break;
	}

	// Try with parent process
	for (int sessionIndex = 0; sessionIndex < sessionCount && pFoundBestSessionControl == NULL; sessionIndex++)
	{
		pSessionControl = pSessions[sessionIndex];
		pSessionControl2 = pSessions2[sessionIndex];

		// Get audio session process id
		DWORD dSessionProcId;
		hr = pSessionControl2->GetProcessId(&dSessionProcId);
		if (FAILED(hr) || dSessionProcId == 0) continue;

		// Return parent process' audio session instead.
		if (dSessionProcId == pProcFamily->ParentId)
		{
			LOG("Found session by matching parent process: %d\n", pProcFamily->ParentId);
			pFoundBestSessionControl = pSessionControl;
			break;
		}
	}

Exit:
	SAFE_RELEASE(pAudioSessionManager);
	SAFE_RELEASE(pSessionEnumerator);

	if (SUCCEEDED(hr))
		*pBestSessionControl = pFoundBestSessionControl;

	if (*pBestSessionControl == NULL)
	{
		LOG_ERR("No session matching PID %d was found\n", dProcId);
		hr = S_FALSE;
	}

	// Release unused sessions
	if (*pBestSessionControl != NULL)
	{
		IAudioSessionControl* pSession;
		for (int sessionIndex = 0; sessionIndex < sessionCount; sessionIndex++)
		{
			pSession = pSessions[sessionIndex];
			if (pSession != *pBestSessionControl)
				SAFE_RELEASE(pSession);
		}
	}

	return hr;
}

/// <summary>
/// Toggle mute the audio session belonging to the given process id, or a
/// process related to it.
/// </summary>
/// <param name="dProcId"></param>
/// <param name="pEndpoint"></param>
/// <returns>S_OK if toggle the mute state. E_FAIL on Windows errors. S_FALSE if a
/// matching audio session was not found. </returns>
HRESULT MuteProcessOnAudioEndpoint(DWORD dProcId, IMMDevice* pEndpoint)
{
	HRESULT hr = E_FAIL;
	BOOL found = FALSE;

	IAudioSessionControl* pSessionControl = NULL;
	IAudioSessionControl2* pSessionControl2 = NULL;
	ISimpleAudioVolume* pSimpleAudioVolume = NULL;

	hr = FindPreferredAudioSession(dProcId, pEndpoint, &pSessionControl);
	EXIT_ON_ERROR(hr, __LINE__);
	if (hr == S_FALSE) // not found
		goto Exit;

	// Get controller v2
	hr = pSessionControl->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&pSessionControl2);
	EXIT_ON_ERROR(hr, __LINE__);

#ifdef _DEBUG
	LOG("Found suitable audio session:\n");
	LogAudioSession(pSessionControl2);
#endif

	// Get audio changer thing
	hr = pSessionControl2->QueryInterface(__uuidof(ISimpleAudioVolume), (void**)&pSimpleAudioVolume);
	EXIT_ON_ERROR(hr, __LINE__);

	// Toggle mute state
	BOOL bMute;
	hr = pSimpleAudioVolume->GetMute(&bMute);
	EXIT_ON_ERROR(hr, __LINE__);

	hr = pSimpleAudioVolume->SetMute(!bMute, &AudioSessionVolumeCtx);
	EXIT_ON_ERROR(hr, __LINE__);

	found = TRUE;

Exit:
	SAFE_RELEASE(pSessionControl);
	SAFE_RELEASE(pSessionControl2);
	SAFE_RELEASE(pSimpleAudioVolume);

	if (!found && SUCCEEDED(hr))
		hr = S_FALSE;
	return hr;
}

HRESULT GlobalMuteProcess(DWORD dProcId)
{
	HRESULT hr;
	HRESULT* hrResults = NULL;

	IMMDeviceEnumerator* pEnumerator = NULL;
	IMMDeviceCollection* pCollection = NULL;
	IMMDevice* pEndpoint = NULL;
	UINT deviceCount = 0;

	// Get audio device enumerator
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
		(void**)&pEnumerator);
	EXIT_ON_ERROR(hr, __LINE__);

	// Get audio device collection
	hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pCollection);
	EXIT_ON_ERROR(hr, __LINE__);

	hr = pCollection->GetCount(&deviceCount);
	EXIT_ON_ERROR(hr, __LINE__);

	hrResults = (HRESULT*)malloc(sizeof(HRESULT*) * deviceCount);

	for (int deviceIndex = 0; deviceIndex < deviceCount; deviceIndex++)
	{
		LOG("--------------\n");
		// Get endpoint
		hr = pCollection->Item(deviceIndex, &pEndpoint);
		if (FAILED(hr))
		{
			LOG_ERR("Failed to get audio endpoint device: %d\n", hr);
			continue;
		}

#ifdef _DEBUG
		LogDeviceEndpoint(pEndpoint);
#endif

		hrResults[deviceIndex] = MuteProcessOnAudioEndpoint(dProcId, pEndpoint);
		SAFE_RELEASE(pEndpoint);
		LOG("--------------\n\n");
	}
Exit:
	SAFE_RELEASE(pEnumerator);
	SAFE_RELEASE(pCollection);

	if (hrResults != NULL)
	{
		for (int i = 0; i < deviceCount; i++)
		{
			if (hrResults[i] == S_OK)
			{
				hr = S_OK;
				break;
			}
		}
	}

	return hr;
}

HRESULT MuteFocusedProcess()
{
	HRESULT hr;
	DWORD dFocusedProcessId;

	// Get focused process id
	hr = GetFocusedProcessId(&dFocusedProcessId);
	EXIT_ON_ERROR(hr, __LINE__);

	hr = GlobalMuteProcess(dFocusedProcessId);
	if (hr == S_FALSE)
	{
		LOG_ERR("Could not mute the process\n");
	}


Exit:
	return hr;
}

int wmain()
{
	HRESULT hr;
	LOG("Entry\n");

	hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	if (FAILED(hr))
	{
		LOG_ERR("Failed to initialize the COM library: %d\n", hr);
		return hr;
	}

#ifdef _DEBUG
	Sleep(1500);
	hr = MuteFocusedProcess();
	std::cin.get();
#else
	ShowWindow(GetConsoleWindow(), SW_HIDE);
	hr = MuteFocusedProcess();
#endif

	return hr;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow) {
	return wmain();
}