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

// https://github.com/microsoft/Windows-classic-samples/blob/d338bb385b1ac47073e3540dbfa810f4dcb12ed8/Samples/Win7Samples/multimedia/mediafoundation/MFPlayer2/AudioSessionVolume.cpp#L30
// {2715279F-4139-4ba0-9CB1-B351F1B58A4A}
static const GUID AudioSessionVolumeCtx = { 0x2715279f, 0x4139, 0x4ba0, { 0x9c, 0xb1, 0xb3, 0x51, 0xf1, 0xb5, 0x8a, 0x4a } };

HRESULT GetFocusedProcessId(DWORD* dpProcId)
{
	HWND wFocusedWind;

	wFocusedWind = GetForegroundWindow();
	if (wFocusedWind == NULL) {
		LOG_ERR("GetForegroundWindow fail");
		return E_FAIL;
	}

	GetWindowThreadProcessId(wFocusedWind, dpProcId);
	if (dpProcId == NULL) {
		LOG_ERR("GetWindowThreadProcessId fail");
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
				LOG("Parent: %d\n", pe.th32ParentProcessID);
			}
			else if (pe.th32ParentProcessID == dProcId) {
				LOG("Child: %d\n", pe.th32ProcessID);
			}
		} while (Process32Next(h, &pe));
	}

	CloseHandle(h);
	return dProcId;
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

HRESULT MuteFocusedProcess()
{
	HRESULT hr = E_FAIL;

	IMMDeviceEnumerator* pEnumerator = NULL;
	IAudioSessionManager2* pAudioSessionManager = NULL;
	IAudioSessionEnumerator* pSessionEnumerator = NULL;
	IMMDeviceCollection* pCollection = NULL;
	IMMDevice* pEndpoint = NULL;
	IAudioEndpointVolume* pInterface = NULL;

	IAudioSessionControl* pSessionControl = NULL;
	IAudioSessionControl2* pSessionControl2 = NULL;
	ISimpleAudioVolume* pSimpleAudioVolume = NULL;

	DWORD dFocusedProcessId;

	hr = CoInitialize(NULL);
	EXIT_ON_ERROR(hr, __LINE__);

	// Get focused process id
	hr = GetFocusedProcessId(&dFocusedProcessId);
	EXIT_ON_ERROR(hr, __LINE__);

	dFocusedProcessId = GetParentProcessId(dFocusedProcessId);

	// Get audio device enumerator
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
		(void**)&pEnumerator);
	EXIT_ON_ERROR(hr, __LINE__);

	// Get audio device collection
	UINT count;
	hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pCollection);
	EXIT_ON_ERROR(hr, __LINE__);
	
	hr = pCollection->GetCount(&count);
	EXIT_ON_ERROR(hr, __LINE__);

	for (int i = 0; i < count; i++)
	{
		// Get endpoint
		hr = pCollection->Item(i, &pEndpoint);
		EXIT_ON_ERROR(hr, __LINE__);

		// Get audio client
		pEndpoint->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pInterface);
		EXIT_ON_ERROR(hr, __LINE__);

		// Get the IAudioSessionManager2 interface
		hr = pEndpoint->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, NULL, (void**)&pAudioSessionManager);
		EXIT_ON_ERROR(hr, __LINE__);

		// Enumerate audio sessions
		hr = pAudioSessionManager->GetSessionEnumerator(&pSessionEnumerator);
		EXIT_ON_ERROR(hr, __LINE__);

		int sessionCount;
		hr = pSessionEnumerator->GetCount(&sessionCount);
		EXIT_ON_ERROR(hr, __LINE__);

		// Looping over the sessions
		for (int i = 0; i < sessionCount; i++) {
			hr = pSessionEnumerator->GetSession(i, &pSessionControl);
			if (FAILED(hr)) continue;

			hr = pSessionControl->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&pSessionControl2);
			if (FAILED(hr)) continue;

			DWORD dProcId;
			hr = pSessionControl2->GetProcessId(&dProcId);
			if (FAILED(hr)) continue;

			LOG("PID: %d\n", dProcId);

			if (dProcId != dFocusedProcessId && !IsProcessRelated(dProcId, dFocusedProcessId))
				continue;

			hr = pSessionControl2->QueryInterface(__uuidof(ISimpleAudioVolume), (void**)&pSimpleAudioVolume);
			EXIT_ON_ERROR(hr, __LINE__);

			BOOL bMute;
			pSimpleAudioVolume->GetMute(&bMute);
			hr = pSimpleAudioVolume->SetMute(!bMute, &AudioSessionVolumeCtx);
			EXIT_ON_ERROR(hr, __LINE__);

			LOG("Ok\n");
			return S_OK;
		}
	}

	// If the code reaches here, we failed to mute the target process.
	LOG_ERR("Not found\n");

Exit:
	SAFE_RELEASE(pEnumerator);
	SAFE_RELEASE(pAudioSessionManager);
	SAFE_RELEASE(pSessionEnumerator);
	SAFE_RELEASE(pCollection);
	SAFE_RELEASE(pEndpoint);
	SAFE_RELEASE(pInterface);
	SAFE_RELEASE(pSessionControl);
	SAFE_RELEASE(pSessionControl2);
	SAFE_RELEASE(pSimpleAudioVolume);

	return hr;
}

int wmain()
{
#ifndef _DEBUG
	ShowWindow(GetConsoleWindow(), SW_HIDE);
#endif

	LOG("Entry\n");
	//Sleep(1500);
	HRESULT hr = MuteFocusedProcess();
	//std::cin.get();
	return hr;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow) {
	return wmain();
}