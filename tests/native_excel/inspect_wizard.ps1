param([long]$WindowHandle)
Add-Type -AssemblyName Accessibility
Add-Type -ReferencedAssemblies Accessibility -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class WizardAccessibility {
    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid,
        [MarshalAs(UnmanagedType.Interface)] out Accessibility.IAccessible result);
    [DllImport("oleacc.dll")]
    public static extern int AccessibleChildren(Accessibility.IAccessible obj, int start, int count,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] object[] children, out int obtained);
}
"@
$iid = [guid]'618736e0-3c3d-11cf-810c-00aa00389b71'
$root = $null
$hr = [WizardAccessibility]::AccessibleObjectFromWindow([intptr]$WindowHandle, [uint32]::MaxValue - 3, [ref]$iid, [ref]$root)
if ($hr -ne 0) { throw "AccessibleObjectFromWindow HRESULT=$hr" }
$script:remaining = 500
function Read-Accessible($obj, $child, [int]$depth) {
    if ($script:remaining-- -le 0 -or $depth -gt 12) { return }
    try { $obj.accName($child) } catch {}
    try { $obj.accDescription($child) } catch {}
    try { $obj.accHelp($child) } catch {}
    try { $obj.accValue($child) } catch {}
    if ($child -ne 0) { return }
    try { $count = [Math]::Min(100, $obj.accChildCount) } catch { return }
    if ($count -le 0) { return }
    $children = [object[]]::new($count)
    $obtained = 0
    [void][WizardAccessibility]::AccessibleChildren($obj, 0, $count, $children, [ref]$obtained)
    for ($i=0; $i -lt $obtained; $i++) {
        if ($children[$i] -is [int]) { Read-Accessible $obj $children[$i] ($depth+1) }
        elseif ($children[$i] -is [Accessibility.IAccessible]) { Read-Accessible $children[$i] 0 ($depth+1) }
    }
}
[Console]::OutputEncoding = [Text.UTF8Encoding]::new()
Read-Accessible $root 0 0
