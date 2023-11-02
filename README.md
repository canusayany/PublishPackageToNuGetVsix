Copy  https://github.com/thxu/PublishPackageToNuGetVsix

做了以下修改:
1.更新了nuget版本,之前的版本可能不稳定.
2.添加了读取文件version,避免了每次都要自己输入
3.读取git最近的提交历史,作为nuget的说明
4.修改了push数据包的逻辑.
