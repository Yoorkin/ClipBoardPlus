using Microsoft.Win32;

/// <summary>
/// 简易程序注册表管理类
/// </summary>
public class AppRegistry
{
    private RegistryKey key;
    private RegistryKey runkey = Registry.LocalMachine.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run");
    private string startup, appname;
    private bool autorun = false;

    /// <summary>
    /// 加载已经注册过的AppRegistry
    /// </summary>
    /// <param name="AppName">注册表中使用的键名</param>
    public AppRegistry(string AppName)
    {
        key = Registry.LocalMachine.CreateSubKey("software\\" + AppName);
        foreach (string key in runkey.GetValueNames())
            if (key == AppName) autorun = true;
        appname = AppName;
    }
    /// <summary>
    /// 加载并注册AppRegistry
    /// </summary>
    /// <param name="AppName">注册表中使用的键名</param>
    /// <param name="StartupPath">启动程序的绝对路径，包括文件名和后缀。建议填AppDomain.CurrentDomain.BaseDirectory/AppName.exe</param>
    public AppRegistry(string AppName, string StartupPath)
    {
        key = Registry.LocalMachine.CreateSubKey("software\\" + AppName);
        foreach (string key in runkey.GetValueNames())
            if (key == AppName) autorun = true;
        appname = AppName;
        startup = StartupPath;
    }


    /// <summary>
    /// 程序是否开机启动
    /// </summary>
    public bool AutoRun
    {
        set
        {
            if (value)
                runkey.SetValue(appname, startup);
            else
                runkey.DeleteValue(appname);

            autorun = value;
        }
        get
        {
            return autorun;
        }
    }

    /// <summary>
    /// 设置键值，如果键值不存在则自动创建
    /// </summary>
    public void Set(string name, object value)
    {
        key.SetValue(name, value);
    }
    /// <summary>
    /// 获取键值
    /// </summary>
    public object Get(string name)
    {
        object obj = key.GetValue(name);
        if (obj.ToString() == "True" || obj.ToString() == "False")
            return obj.ToString() == "True" ? true : false;
        else
            return obj;
    }
    /// <summary>
    /// 获取键值，如果键值不存在则返回DefaultValue
    /// </summary>
    public object Get(string name, object DefaultValue)
    {
        return key.GetValue(name, DefaultValue);
    }
    public void Remove(string name)
    {
        key.DeleteValue(name);
    }

    public void DeleteAppRegInfo()
    {
        Registry.LocalMachine.CreateSubKey("SoftWare").DeleteSubKey(appname);
    }
}

