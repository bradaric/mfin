//! Windows 11 shell extension for mfin — IExplorerCommand implementation.
//!
//! This DLL registers a context menu item "Extract Tables (mfin)" that appears
//! in the Windows 11 modern right-click menu for PDF files. It launches
//! mfin.exe with the selected file path.
//!
//! Build: cargo build --release --target x86_64-pc-windows-msvc
//! Output: target/x86_64-pc-windows-msvc/release/mfin_shell.dll

use std::sync::atomic::{AtomicU32, Ordering};

use windows::core::*;
use windows::Win32::Foundation::*;
use windows::Win32::System::Com::*;
use windows::Win32::System::Ole::*;
use windows::Win32::UI::Shell::*;
use windows::Win32::UI::WindowsAndMessaging::*;

/// Must match SHELL_EXT_CLSID in windows_gui.py.
const CLSID_MFIN_EXTRACT: GUID = GUID::from_u128(0xe7a20a14_3b78_4d9a_9c12_5f4b8a3e6d01);

static DLL_REF_COUNT: AtomicU32 = AtomicU32::new(0);
static mut HINSTANCE_DLL: Option<HMODULE> = None;

// ---- DLL entry points ----

#[no_mangle]
unsafe extern "system" fn DllMain(hinst: HMODULE, reason: u32, _reserved: *mut ()) -> BOOL {
    if reason == 1 {
        // DLL_PROCESS_ATTACH
        HINSTANCE_DLL = Some(hinst);
    }
    TRUE
}

#[no_mangle]
unsafe extern "system" fn DllGetClassObject(
    rclsid: *const GUID,
    riid: *const GUID,
    ppv: *mut *mut core::ffi::c_void,
) -> HRESULT {
    if ppv.is_null() {
        return E_POINTER;
    }
    *ppv = std::ptr::null_mut();

    if *rclsid != CLSID_MFIN_EXTRACT {
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    let factory: IClassFactory = MfinClassFactory.into();
    match factory.query(&*riid, ppv) {
        Ok(()) => S_OK,
        Err(e) => e.code(),
    }
}

#[no_mangle]
extern "system" fn DllCanUnloadNow() -> HRESULT {
    if DLL_REF_COUNT.load(Ordering::SeqCst) == 0 {
        S_OK
    } else {
        S_FALSE
    }
}

// ---- Helpers ----

/// Allocate a COM-compatible wide string (caller frees with CoTaskMemFree).
fn alloc_com_str(s: &str) -> PWSTR {
    let wide: Vec<u16> = s.encode_utf16().chain(std::iter::once(0)).collect();
    let byte_len = wide.len() * 2;
    unsafe {
        let ptr = CoTaskMemAlloc(byte_len) as *mut u16;
        if !ptr.is_null() {
            std::ptr::copy_nonoverlapping(wide.as_ptr(), ptr, wide.len());
        }
        PWSTR(ptr)
    }
}

/// Find mfin.exe next to this DLL.
fn find_mfin_exe() -> String {
    unsafe {
        let hinst = match HINSTANCE_DLL {
            Some(h) => h,
            None => return String::new(),
        };
        let mut buf = [0u16; 260];
        let len = GetModuleFileNameW(Some(hinst), &mut buf);
        if len == 0 {
            return String::new();
        }
        let dll_path = String::from_utf16_lossy(&buf[..len as usize]);
        if let Some(dir) = std::path::Path::new(&dll_path).parent() {
            let exe = dir.join("mfin.exe");
            if exe.exists() {
                return exe.to_string_lossy().into_owned();
            }
        }
        String::new()
    }
}

/// Launch mfin.exe with the given PDF path.
fn launch_mfin(pdf_path: &str) {
    let exe = find_mfin_exe();
    if exe.is_empty() {
        return;
    }
    unsafe {
        let exe_w: HSTRING = exe.into();
        let params: HSTRING = format!("\"{}\"", pdf_path).into();
        ShellExecuteW(None, w!("open"), &exe_w, &params, None, SW_SHOWNORMAL);
    }
}

// ---- IClassFactory ----

#[implement(IClassFactory)]
struct MfinClassFactory;

impl IClassFactory_Impl for MfinClassFactory_Impl {
    fn CreateInstance(
        &self,
        _punkouter: Option<&IUnknown>,
        riid: *const GUID,
        ppvobject: *mut *mut core::ffi::c_void,
    ) -> Result<()> {
        unsafe {
            if ppvobject.is_null() {
                return Err(E_POINTER.into());
            }
            *ppvobject = std::ptr::null_mut();

            let cmd: IExplorerCommand = MfinExplorerCommand.into();
            match cmd.query(&*riid, ppvobject) {
                Ok(()) => Ok(()),
                Err(e) => Err(e),
            }
        }
    }

    fn LockServer(&self, flock: BOOL) -> Result<()> {
        if flock.as_bool() {
            DLL_REF_COUNT.fetch_add(1, Ordering::SeqCst);
        } else {
            DLL_REF_COUNT.fetch_sub(1, Ordering::SeqCst);
        }
        Ok(())
    }
}

// ---- IExplorerCommand ----

#[implement(IExplorerCommand)]
struct MfinExplorerCommand;

impl IExplorerCommand_Impl for MfinExplorerCommand_Impl {
    fn GetTitle(&self, _psiitemarray: Option<&IShellItemArray>) -> Result<PWSTR> {
        Ok(alloc_com_str("Extract Tables (mfin)"))
    }

    fn GetIcon(&self, _psiitemarray: Option<&IShellItemArray>) -> Result<PWSTR> {
        let exe = find_mfin_exe();
        if exe.is_empty() {
            Ok(PWSTR::null())
        } else {
            Ok(alloc_com_str(&exe))
        }
    }

    fn GetToolTip(&self, _psiitemarray: Option<&IShellItemArray>) -> Result<PWSTR> {
        Ok(alloc_com_str("Extract tables from this PDF using mfin"))
    }

    fn GetCanonicalName(&self) -> Result<GUID> {
        Ok(CLSID_MFIN_EXTRACT)
    }

    fn GetState(
        &self,
        _psiitemarray: Option<&IShellItemArray>,
        _foktobeslow: BOOL,
    ) -> Result<EXPCMDSTATE> {
        Ok(ECS_ENABLED)
    }

    fn Invoke(
        &self,
        psiitemarray: Option<&IShellItemArray>,
        _pbc: Option<&IBindCtx>,
    ) -> Result<()> {
        let items = psiitemarray.ok_or(E_INVALIDARG)?;

        unsafe {
            let count = items.GetCount()?;
            for i in 0..count {
                let item = items.GetItemAt(i)?;
                let path = item.GetDisplayName(SIGDN_FILESYSPATH)?;
                let path_str = path.to_string()?;
                CoTaskMemFree(Some(path.0 as *mut _));

                launch_mfin(&path_str);
            }
        }
        Ok(())
    }

    fn GetFlags(&self) -> Result<EXPCMDFLAGS> {
        Ok(ECF_DEFAULT)
    }

    fn EnumSubCommands(&self) -> Result<IEnumExplorerCommand> {
        Err(E_NOTIMPL.into())
    }
}
