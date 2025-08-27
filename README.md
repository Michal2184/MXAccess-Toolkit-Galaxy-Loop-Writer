# MXAccess Galaxy Loop Writer (C++)

A minimal, production-safe C++ example that **writes to an AVEVA (Wonderware) Galaxy attribute in a loop** using the **MX Access Toolkit**.

This sample:

- Registers a client with MX Access
- Adds the target attribute (default: `Car.speed`)
- Puts the item on **Advise** (required before writes)
- Runs a loop that **writes every _N_ ms**
- **Pumps the STA message queue** so COM can deliver async completions (critical for live updates)
- Cleans up (`UnAdvise → RemoveItem → Unregister`)

> ⚠️ Build as **Win32/x86**. The MX Access COM server is 32-bit; a 64-bit build will fail with `0x80040154 (Class not registered)`.

---

## Table of contents

- [Prerequisites](#prerequisites)  
- [Project setup (Visual Studio)](#project-setup-visual-studio)  
- [Build & run](#build--run)  
- [Usage & configuration](#usage--configuration)  
- [How it works](#how-it-works)  
- [Optional enhancements](#optional-enhancements)  
- [Troubleshooting](#troubleshooting)  
- [FAQ](#faq)  
- [License](#license)  
- [Acknowledgements](#acknowledgements)

---

## Prerequisites

- Windows 10/11
- Visual Studio 2019/2022 with Desktop C++ workload
- **AVEVA / Wonderware MX Access Toolkit** installed  
  (provides `MxAccess32.tlb` and the `LMXProxyServer` COM server)
- A running Galaxy and a **writable attribute** to target (e.g., `Car.speed`)

---

## Project setup (Visual Studio)

1. **Create** → *Console App* (C++)  
2. **Platform** → **Win32 / x86** (not x64)  
3. **Language standard** → C++17 or newer (optional)  
4. **Add source** → drop in `MXAccess-Simple-Loop-Write.cpp` from this repo  
5. Make sure the `#import` path matches your install (adjust if needed):
   ```cpp
   #import "C:\\Program Files (x86)\\ArchestrA\\Framework\\bin\\MxAccess32.tlb" \
     no_namespace raw_interfaces_only rename("EOF","MX_EOF")
   ```

---

## Build & run

1. Build **Win32/Debug** (or Release).  
2. Ensure the Galaxy/platform is deployed and on scan and have installed MXAccess Toolkit. 
3. Run the app. Example console output:
   ```
   Register OK (hLMX=1)
   AddItem OK (hItem=1)
   Advise (User) OK

   Write 0
   Write 1
   Write 2
   ...
   ```

You should see the attribute updating in your monitoring tools.

---

## Usage & configuration

Key knobs in the sample:

```cpp
DWORD intervalMs = 50;   // write period (ms)
int   startValue = 0;    // starting counter

// Target attribute (edit to match your Galaxy):
hr = lmx2->AddItem(hLMX, _bstr_t(L"Car.speed"), &hItem);

// Current write uses a double value:
SetDouble(vVal, static_cast<double>(counter));
// (vTime is prepared if you later switch to Write2)
```

- **Change target:** replace `"Car.speed"` with your attribute path.  
- **Write type:** use `SetInteger`, `SetDouble`, or `SetString` helpers to send the type you need.  
- **Loop exit:** the sample runs indefinitely; add a keypress check if you want to stop gracefully or use Ctrl-C

---

## How it works

- **COM init (STA):**  
  `CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);` creates a Single-Threaded Apartment.

- **Connect to MX Access:**  
  `LMXProxyServer` is created, then `Register(...)` returns a session handle (`hLMX`).

- **Add & Advise:**  
  `AddItem(...)` adds the attribute; `Advise(...)` is **required before writes**.

- **Loop write + STA pump:**  
  Every `intervalMs`:
  - A `VARIANT` is prepared and `Write(hLMX, hItem, vVal, userID)` is called.
  - The loop runs a **message pump**:
    ```cpp
    while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
      TranslateMessage(&msg);
      DispatchMessage(&msg);
    }
    ```
    In an STA, COM delivers callbacks through the thread’s message queue.  
    Pumping keeps async write completions flowing so values update live (instead of “pending” until exit).

- **Cleanup:**  
  `UnAdvise → RemoveItem → Unregister → CoUninitialize`.

---

## Optional enhancements

You can evolve this minimal sample with the following (left out by default to keep it simple):

### Supervisory advise (no user login required for Secured/Verified attributes)
```cpp
CComQIPtr<ILMXProxyServer4> lmx4(lmx2);
if (lmx4) {
  hr = lmx4->AdviseSupervisory(hLMX, hItem);
} else {
  hr = lmx2->Advise(hLMX, hItem);
}
```

### Write2 (value + timestamp)
Provide a VT_DATE timestamp and call `Write2` (requires `ILMXProxyServer4`):
```cpp
VARIANT vVal, vTime;
SetDouble(vVal, value);
SetNowAsVariantTime(vTime); // local time; or build high-res/UTC if preferred
lmx4->Write2(hLMX, hItem, vVal, vTime, userID);
```

### Event-gated writes
Connect `_ILMXProxyServerEvents::OnWriteComplete` and only issue the next write **after** the previous one completes for maximum reliability at high rates.

---

## Troubleshooting

**`0x80040154 (Class not registered)` on `CoCreateInstance`**  
You built **x64** or MX Access isn’t registered. Switch to **Win32/x86** and ensure the Toolkit is installed.

**No updates until the app exits**  
Ensure the **message pump** runs inside the loop. Consider event-gated writes if you’re pushing very fast.

**`AddItem` or `Advise` fails**  
- Verify the attribute path is correct and writeable.  
- Check the platform is online and healthy.

**Writes ignored on Secured/Verified attributes**  
Use **`AdviseSupervisory(...)`** (recommended) or authenticate and use `WriteSecured(...)`.

---

## FAQ

**Q: Why x86 only?**  
A: MX Access is a 32-bit COM server. Your client must be 32-bit to instantiate it.

**Q: Do I need the timestamp?**  
A: No. `Write()` is fine for many cases. `Write2()` (value + timestamp) can be useful for precise VTQ handling.

**Q: Why the message pump if I’m not handling events?**  
A: In an STA, COM completions still rely on the thread’s message queue. Pumping avoids “pending till exit” behavior.

---

## License

MIT — see `LICENSE` (add one if you plan to distribute).

---

## Acknowledgements

“AVEVA”, “Wonderware”, “Galaxy”, and “MX Access Toolkit” are trademarks of their respective owners. This project is a community example and is not affiliated with or endorsed by AVEVA.

