
# 首先安装PYRFC：
官方文档：https://github.com/sap/pyrfc/

##  安装 NWRFC SDK 
1. NWRFC SDK安装说明：
-     首先配置 SAPNWRFC_HOME 环境变量，lib 文件夹应该是 SAPNWRFC_HOME 环境变量指向文件夹直接下层文件夹

-     接下来配置 Path 环境变量，将 nwrfcsdk 文件夹下面的 lib 和 bin 文件夹加入到 Path 环境变量：
2. 例如：
-     安装目录：C:\nwrfcsdk\nwrfcsdk：SAPNWRFC_HOME = C:\nwrfcsdk\nwrfcsdk
-     path： 增加 C:\nwrfcsdk\nwrfcsdk\bin，C:\nwrfcsdk\nwrfcsdk\lib

## 安装 pyrfc 
1. pip install pyrfc  直接安装 pyrfc
2. 或者：https://github.com/SAP/PyRFC/tags ,下载对应的 whl版本，然后  pip install pyrfc-XXX-XXX-win_AMD64.whl 安装

# 安装azsaprfc
pip install azsaprfc
