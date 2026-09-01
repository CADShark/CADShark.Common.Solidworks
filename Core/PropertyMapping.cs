using SolidWorks.Interop.swconst;

namespace CADShark.Common.SolidWorks.Core
{
    /// <summary>
    /// Pure managed mapping of SOLIDWORKS custom-property COM result codes to
    /// documented application semantics. This type contains only the code that
    /// decides how a raw COM result is interpreted; it performs no COM calls, so
    /// the rules can be unit-tested in isolation from a live SldWorks instance.
    ///
    /// A4 HOST/ADAPTER BOUNDARY: <see cref="SwPropertyManager"/> is the COM-owning
    /// choke point for custom-property reads and writes. Its result-code
    /// interpretation is delegated here so the decision logic is a single,
    /// testable seam instead of being inlined in every overload.
    /// </summary>
    public static class PropertyMapping
    {
        /// <summary>
        /// Maps a <c>ICustomPropertyManager::Get6</c> result code to the resolved
        /// property value. Only <see cref="swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue"/>
        /// returns the value; every other outcome yields an empty string,
        /// preserving the existing <see cref="SwPropertyManager.GetProperty"/> contract.
        /// </summary>
        public static string MapGetResult(int resultCode, string resolvedValue)
        {
            return resultCode == (int)swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue
                ? resolvedValue
                : "";
        }

        /// <summary>
        /// Maps an <c>ICustomPropertyManager::Add3</c> result code to success.
        /// Only <see cref="swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged"/>
        /// is treated as success, preserving the existing <see cref="SwPropertyManager.SetProperty"/>
        /// contract (which surfaces a message box on any other outcome).
        /// </summary>
        public static bool IsAddSuccessful(int resultCode)
        {
            return resultCode == (int)swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged;
        }
    }
}
