//------------------------------------------------------------------------------
// MXAccess C++ Simple loop write speed test (STA message pump, no readback)
// This example writes an incrementing integer value to the Car object 
// with attribute speed - "Car.speed"
//------------------------------------------------------------------------------

#include <windows.h>
#include <comdef.h>
#include <atlbase.h>
#include <iostream>
#include <string>
// install MX Access toolkit to get MxAccess32.tlb
#import "C:\\Program Files (x86)\\ArchestrA\\Framework\\bin\\MxAccess32.tlb" no_namespace raw_interfaces_only rename("EOF","MX_EOF")

typedef long ITEMHANDLE;

// Helpers for data types: integers , doubles, strings and VT_DATE
static inline void SetDouble(VARIANT& v, double d) {
    VariantInit(&v);
    v.vt = VT_R8;
    v.dblVal = d;
}

static inline void SetInteger(VARIANT& v, long i) {
    VariantInit(&v);
    v.vt = VT_I4;
    v.lVal = i;
}

static inline void SetString(VARIANT& v, const std::wstring& s) {
    VariantInit(&v);
    v.vt = VT_BSTR;
    v.bstrVal = SysAllocString(s.c_str());
}

static inline void SetNowAsVariantTime(VARIANT& v) {
    VariantInit(&v);
    v.vt = VT_DATE;
    SYSTEMTIME st; GetLocalTime(&st);
    SystemTimeToVariantTime(&st, &v.date);
}

// Simple STA message pump: allows COM deliver callbacks/completions.
static inline void PumpMessagesOnce() {
    MSG msg;
    while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

int main() {
	// timing and loop control
    DWORD intervalMs = 50;
    int   startValue = 0;

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) { 
        std::cout << "CoInitializeEx" << hr; 
        return 1; 
    }

    // Create LMX server
    CComPtr<ILMXProxyServer2> lmx2;
    hr = lmx2.CoCreateInstance(__uuidof(LMXProxyServer), nullptr, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER);
    if (FAILED(hr) || !lmx2) {
        std::cout << "CoCreateInstance(LMXProxyServer" << hr;
        CoUninitialize(); 
        return 1;
    }
    CComQIPtr<ILMXProxyServer4> lmx4(lmx2);

	// Register client
    long hLMX = 0;
    hr = lmx2->Register(_bstr_t(L"LoopSpeedWriter"), &hLMX);
    if (FAILED(hr) || !hLMX) { 
    std::cout << "Register" << hr;;
        CoUninitialize(); 
        return 1; 
    }
    std::wcout << L"Register OK (hLMX=" << hLMX << L")\n";

	// Add item - Car.speed
    ITEMHANDLE hItem = 0;
    hr = lmx2->AddItem(hLMX, _bstr_t(L"Car.speed"), &hItem);
    if (FAILED(hr) || !hItem) {
        std::cout << "AddItem(Car.speed)" << hr;
        lmx2->Unregister(hLMX);
        CoUninitialize(); return 1;
    }
    std::wcout << "AddItem OK " << hItem << L")\n";

	// Advise
    hr = lmx2->Advise(hLMX, hItem);
    if (SUCCEEDED(hr)) std::wcout << L"Advise (User) OK\n";
    if (FAILED(hr)) {
        std::cout << "Advise/AdviseSupervisory" << hr;
        lmx2->RemoveItem(hLMX, hItem);
        lmx2->Unregister(hLMX);
        CoUninitialize(); 
        return 1;
    }

    std::wcout << std::endl;

    // Loop config
    long userID = 0; // no security
    int  counter = startValue;
    VARIANT vVal, vTime; 
    VariantInit(&vVal); 
    VariantInit(&vTime);
    DWORD lastTick = GetTickCount64();

    while (true) {
        // timing
        DWORD now = GetTickCount64();
        DWORD elapsed = now - lastTick;
        if (elapsed < intervalMs) {
            Sleep(intervalMs - elapsed);
            now = GetTickCount64();
        }
        lastTick = now;

        //message pump to service COM (writes will not work if not triggered)
        PumpMessagesOnce();
		// variant for value and time
        VariantClear(&vVal);
        VariantClear(&vTime);
        SetDouble(vVal, static_cast<double>(counter));
        SetNowAsVariantTime(vTime);

		// write
        hr = lmx2->Write(hLMX, hItem, vVal, userID);
        if (FAILED(hr)) {
            std::cout << "Write failed: " << hr;
        }
        else {
            std::cout << "Write " << counter << "\n";
        }

        ++counter;
		// reset counter to avoid overflow
        if (counter >= 100000000) counter = 0;
    }

    // Tidy up
    std::cout << "\nStopping… \n";
    VariantClear(&vVal);
    VariantClear(&vTime);
    HRESULT hr2 = lmx2->UnAdvise(hLMX, hItem);
    hr2 = lmx2->RemoveItem(hLMX, hItem);
    lmx2->Unregister(hLMX);
    CoUninitialize();
    std::cout << "Done.\n";
    return 0;
}
