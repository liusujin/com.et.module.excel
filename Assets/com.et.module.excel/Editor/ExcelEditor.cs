using ETModel;
using UnityEditor;

namespace ETEditor
{
    public class ExcelEditor : Editor
    {
        [MenuItem("Tools/导出配置/客户端")]
        private static void ExportClientConfig()
        {
            ExcelHelper.IsClient = true;
            ExcelHelper.ExportAll(ExcelHelper.ClientConfigPath);

            ExcelHelper.ExportAllClass(@"./Assets/Model/Config/", "namespace ETModel\n{\n");
            ExcelHelper.ExportAllClass(@"./Assets/Hotfix/Config/", "using ETModel;\n\nnamespace ETHotfix\n{\n");

            AssetDatabase.Refresh();
        }

        [MenuItem("Tools/导出配置/服务端")]
        private static void ExportServerConfig()
        {
            ExcelHelper.IsClient = false;
            ExcelHelper.ExportAll(ExcelHelper.ServerConfigPath);

            ExcelHelper.ExportAllClass(@"../Server/Model/Module/Demo/Config", "namespace ETModel\n{\n");

            AssetDatabase.Refresh();
        }
    }
}