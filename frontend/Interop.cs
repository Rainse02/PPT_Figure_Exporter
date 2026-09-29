using System;
using System.Runtime.InteropServices;

namespace PPTFigureExporter
{
    // Office and Extensibility contracts are embedded from Microsoft interop assemblies.
    [ComVisible(true), Guid("93D57BE4-929F-4B0B-8B5F-66C11CF42208")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonCallbacks
    {
        [DispId(1)] void OnRibbonLoad(object ribbon);
        [DispId(2)] string GetSelectionLabel(object control);
        [DispId(3)] string GetFormat(object control);
        [DispId(4)] string GetDpi(object control);
        [DispId(5)] bool GetRasterEnabled(object control);
        [DispId(6)] bool GetTransparencyEnabled(object control);
        [DispId(7)] bool GetTransparent(object control);
        [DispId(8)] string GetMode(object control);
        [DispId(9)] bool GetExportEnabled(object control);
        [DispId(10)] void OnFormat(object control, string selectedId, int selectedIndex);
        [DispId(11)] void OnDpi(object control, string value);
        [DispId(12)] void OnTransparent(object control, bool value);
        [DispId(13)] void OnMode(object control, string selectedId, int selectedIndex);
        [DispId(14)] void OnSettings(object control);
        [DispId(15)] void OnExport(object control);
    }
}
