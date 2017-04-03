echo on
set projectFolderPath = %~1
set nugetPackageFolderPath = %~2

pushd %~1
%~2\nuget pack