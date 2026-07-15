[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$excel = [string](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe').'(default)'
if (-not (Test-Path -LiteralPath $excel -PathType Leaf)) { throw 'Registered Excel executable was not found' }

$source = @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CT = System.Runtime.InteropServices.ComTypes;
public static class ExcelRtdTypeLibAudit {
  [DllImport("oleaut32.dll", CharSet=CharSet.Unicode, PreserveSig=false)]
  static extern void LoadTypeLibEx(string file, int regkind, out CT.ITypeLib lib);
  static string TypeName(CT.ITypeInfo owner, CT.TYPEDESC type) {
    int raw=(int)type.vt, value=raw & 0x0fff;
    if(value==26 || value==27) {
      CT.TYPEDESC inner=(CT.TYPEDESC)Marshal.PtrToStructure(type.lpValue,typeof(CT.TYPEDESC));
      return (value==26?"PTR<":"SAFEARRAY<")+TypeName(owner,inner)+">";
    }
    if(value==29) {
      CT.ITypeInfo reference; owner.GetRefTypeInfo(unchecked((int)type.lpValue.ToInt64()),out reference);
      string name,doc,help; int context; reference.GetDocumentation(-1,out name,out doc,out context,out help);
      return "USER("+name+")";
    }
    return ((VarEnum)raw).ToString();
  }
  static void AddVtable(CT.ITypeInfo dispatch, List<string> output) {
    int href; dispatch.GetRefTypeOfImplType(-1,out href); CT.ITypeInfo type; dispatch.GetRefTypeInfo(href,out type);
    string interfaceName,doc,help; int context; type.GetDocumentation(-1,out interfaceName,out doc,out context,out help);
    IntPtr attrPointer; type.GetTypeAttr(out attrPointer); CT.TYPEATTR attr=(CT.TYPEATTR)Marshal.PtrToStructure(attrPointer,typeof(CT.TYPEATTR));
    output.Add("INTERFACE|"+interfaceName+"|"+attr.guid.ToString("B").ToLowerInvariant()+"|"+attr.cFuncs);
    int elementSize=Marshal.SizeOf(typeof(CT.ELEMDESC));
    for(int slot=0;slot<attr.cFuncs;slot++) {
      IntPtr functionPointer; type.GetFuncDesc(slot,out functionPointer); CT.FUNCDESC function=(CT.FUNCDESC)Marshal.PtrToStructure(functionPointer,typeof(CT.FUNCDESC));
      string[] names=new string[Math.Max(1,function.cParams+1)]; int got; type.GetNames(function.memid,names,names.Length,out got);
      string line="METHOD|"+interfaceName+"|"+slot+"|"+function.memid+"|"+function.callconv+"|"+TypeName(type,function.elemdescFunc.tdesc)+"|"+names[0];
      for(int parameter=0;parameter<function.cParams;parameter++) {
        CT.ELEMDESC element=(CT.ELEMDESC)Marshal.PtrToStructure(IntPtr.Add(function.lprgelemdescParam,parameter*elementSize),typeof(CT.ELEMDESC));
        line+="|"+TypeName(type,element.tdesc)+":0x"+((int)element.desc.paramdesc.wParamFlags).ToString("x");
      }
      output.Add(line); type.ReleaseFuncDesc(functionPointer);
    }
    type.ReleaseTypeAttr(attrPointer);
  }
  public static string[] Inspect(string path) {
    CT.ITypeLib library; LoadTypeLibEx(path,2,out library); List<string> output=new List<string>();
    IntPtr attrPointer; library.GetLibAttr(out attrPointer); CT.TYPELIBATTR attr=(CT.TYPELIBATTR)Marshal.PtrToStructure(attrPointer,typeof(CT.TYPELIBATTR));
    output.Add("LIB|"+attr.guid.ToString("B").ToLowerInvariant()+"|"+attr.wMajorVerNum+"."+attr.wMinorVerNum+"|"+attr.lcid+"|"+attr.syskind); library.ReleaseTLibAttr(attrPointer);
    for(int index=0;index<library.GetTypeInfoCount();index++) {
      CT.ITypeInfo type; library.GetTypeInfo(index,out type); string name,doc,help; int context; type.GetDocumentation(-1,out name,out doc,out context,out help);
      if(name=="IRtdServer" || name=="IRTDUpdateEvent") AddVtable(type,output);
    }
    return output.ToArray();
  }
}
'@
Add-Type -TypeDefinition $source -Language CSharp
$actual = @([ExcelRtdTypeLibAudit]::Inspect($excel))
$expected = @(
    'LIB|{00020813-0000-0000-c000-000000000046}|1.9|0|SYS_WIN32',
    'INTERFACE|IRTDUpdateEvent|{a43788c1-d91b-11d3-8f39-00c04f3651b8}|4',
    'METHOD|IRTDUpdateEvent|0|10|CC_STDCALL|VT_HRESULT|UpdateNotify',
    'METHOD|IRTDUpdateEvent|1|11|CC_STDCALL|VT_HRESULT|HeartbeatInterval|PTR<VT_I4>:0xa',
    'METHOD|IRTDUpdateEvent|2|11|CC_STDCALL|VT_HRESULT|HeartbeatInterval|VT_I4:0x1',
    'METHOD|IRTDUpdateEvent|3|12|CC_STDCALL|VT_HRESULT|Disconnect',
    'INTERFACE|IRtdServer|{ec0e6191-db51-11d3-8f3e-00c04f3651b8}|6',
    'METHOD|IRtdServer|0|10|CC_STDCALL|VT_HRESULT|ServerStart|PTR<USER(IRTDUpdateEvent)>:0x1|PTR<VT_I4>:0xa',
    'METHOD|IRtdServer|1|11|CC_STDCALL|VT_HRESULT|ConnectData|VT_I4:0x1|PTR<SAFEARRAY<VT_VARIANT>>:0x1|PTR<VT_BOOL>:0x3|PTR<VT_VARIANT>:0xa',
    'METHOD|IRtdServer|2|12|CC_STDCALL|VT_HRESULT|RefreshData|PTR<VT_I4>:0x3|PTR<SAFEARRAY<VT_VARIANT>>:0xa',
    'METHOD|IRtdServer|3|13|CC_STDCALL|VT_HRESULT|DisconnectData|VT_I4:0x1',
    'METHOD|IRtdServer|4|14|CC_STDCALL|VT_HRESULT|Heartbeat|PTR<VT_I4>:0xa',
    'METHOD|IRtdServer|5|15|CC_STDCALL|VT_HRESULT|ServerTerminate'
)
$missing = @($expected | Where-Object { $_ -notin $actual })
$unexpected = @($actual | Where-Object { $_ -notin $expected })
if ($missing.Count -ne 0 -or $unexpected.Count -ne 0) {
    throw "Excel RTD type-library ABI differs; missing=$($missing -join ';') unexpected=$($unexpected -join ';')"
}
[PSCustomObject]@{
    status = 'verified'
    type_library = '{00020813-0000-0000-C000-000000000046}'
    version = '1.9'
    lcid = 0
    platform = 'SYS_WIN32 automation type library registered for Win64 Excel'
    excel_file_version = (Get-Item -LiteralPath $excel).VersionInfo.FileVersion
    definitions = $actual
} | ConvertTo-Json -Depth 4
