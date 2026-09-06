# macOS 导表工具

1. 将 `.xlsx` 或 `.xlsm` 配置表放在工具目录（`ExcelToGameConfig_mac.command` 同级）。
2. 确保 macOS 已安装 Python 3。
3. 双击 `ExcelToGameConfig_mac.command`。首次运行会自动创建工具目录下的 `.venv` 并安装 `openpyxl`，后续运行会直接复用该环境。导出结果会写入同级的 `Config/` 目录，生成：
   - `game_config.bin`
   - `game_config.json`
   - `config_version.json`

也可以在终端执行并指定目录或版本号：

```sh
./ExcelToGameConfig_mac.command --input-dir /path/to/excel --output-dir /path/to/Config --version 1.0.0
```

macOS 入口复用 `ExcelToGameConfig.py` 的导出实现，保证与 Windows 工具生成相同的二进制格式。

## 生成独立 App

如果希望目标机器完全不安装 Python，双击 `build_mac_app.command` 构建一次。构建过程会自动准备构建环境，生成 `dist/ExcelToGameConfig_mac.app`。将这个 App 和 Excel 配置表放在同一目录后，直接双击 App 即可运行。

该 App 会包含 Python 和 `openpyxl`，首次构建需要联网下载 PyInstaller；使用生成的 App 时不再需要 Python 或网络。App 按构建机器架构生成：Apple Silicon 和 Intel Mac 需要分别构建。
