前些年做云盘产品的时候，一个很核心的功能就是 Office 文件预览，当时还没有使用 .NET Core ,程序部署在 Windows Server 服务器上，文件预览的方案采用了微软的 OWA 。

目前在做的零代码产品中的表单附件控件，同样面临着 Office 文件预览的问题，现在技术栈采用了 .NET Core ，并使用容器化部署，自然就抛弃了 OWA 的方案。

本文简单介绍下 OWA 的替代方案。

## 思路

1、在表单的附件控件上传 Office 文件后，存储到 MongoDB 中，并发消息给文件转换程序；

2、文件转换程序从 MongoDB 获取 Office 文件，通过 Libreoffice 转换为 PDF 文件；

3、将 PDF 文件存储到 MongoDB 中，并将 PDF 文件在 MongoDB 中的 FileID 存储到平台和原始文件进行关联；

4、在表单中点击文件预览时使用关联的 PDF 的文件 ID 从 MongoDB 中获取 PDF 文件进行展示。

## 准备

1、创建一个 .NET Core 的控制台程序用来做文件的转换；

2、下载 Libreoffice 安装包、Libreoffice 中文语言包、jdk1.8 安装包 、中文字体包，这些文件我放在云盘了，可以访问这个链接下载：https://pan.baidu.com/s/131lLewbCvGDGLlZzYdSYNA 提取码: 5aas

![](https://cdn.jsdelivr.net/gh/oec2003/hblog-images/img/202208281754206.png)

3、搭建一台 centos 虚拟机，并准备好 docker 环境；

## 版本

- .NET Core：3.1
- CentOS：7.6
- Docker：
- Liberoffice：7.3.5
- RabbitMQ：3.8.2
- MongoDB：5.0

## 开始

### 编写控制台程序进行文件转换

1、创建一个名为 OfficeToPdf 的 .NET Core 控制台程序，在 Main 方法中对消息队列进行监听；

```
static void Main(string[] args)  
{  
    try  
    {  
        var mqManager = new MQManager(new MQConfig  
        {  
            AutomaticRecoveryEnabled = true,  
            HeartBeat = 60,  
            NetworkRecoveryInterval = new TimeSpan(60),  
            Host = EnvironmentHelper.GetEnvValue("MQHostName"),  
            UserName = EnvironmentHelper.GetEnvValue("MQUserName"),  
            Password = EnvironmentHelper.GetEnvValue("MQPassword"),  
            Port = EnvironmentHelper.GetEnvValue("MQPort")  
        });        
        if (mqManager.Connected)  
        {            
	        _logger.Log(LogLevel.Info, "RabbitMQ连接成功。");  
            _logger.Log(LogLevel.Info, "RabbitMQ消息接收中...");  
  
            mqManager.Subscribe<PowerPointConvertMessage>(Convert);  
            mqManager.Subscribe<WordConvertMessage>(Convert);  
            mqManager.Subscribe<ExcelConvertMessage>(Convert);  
        }        
        else  
        {  
            _logger.Warn("RabbitMQ连接初始化失败,请检查连接。");  
            Console.ReadLine();  
        }    
    }catch(Exception ex)  
    {        
	    _logger.Error(ex.Message);  
    }
}
```

2、在 Convert 方法中对消息进行处理，首先根据消息的中的文件 ID 获取文件：

```
Stream sourceStream = fileOperation.GetFile(officeMessage.FileInfo.FileId);  
if(sourceStream == null)  
{  
    logger.Log(LogLevel.Error, $"文件ID：{officeMessage.FileInfo.FileId}，不存在");  
}  
  
string filename = officeMessage.FileInfo.FileId;  
string extension = System.IO.Path.GetExtension(officeMessage.FileInfo.FileName);  
  
sourcePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), filename + extension);  
destPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), string.Format("{0}.pdf", filename));  
  
logger.Log(LogLevel.Info, $"文件原路径：{sourcePath}");  
logger.Log(LogLevel.Info, $"文件目标路径：{destPath}");  
if (extension != null && (extension.Equals(".xlsx",StringComparison.OrdinalIgnoreCase) ||   
                          extension.Equals(".xls", StringComparison.OrdinalIgnoreCase)))  
{  
    if (!SetExcelScale(sourceStream, sourcePath))  
        return false;  
}  
else  
{  
    byte[] sourceBuffer = new Byte[sourceStream.Length];  
    sourceStream.Read(sourceBuffer, 0, sourceBuffer.Length);  
    sourceStream.Seek(0, SeekOrigin.Begin);  
    if (!SaveToFile(sourceBuffer, sourcePath))  
        return false;  
}
```

3、启用 LibreOffice 进行文件转换：

```
var psi = new ProcessStartInfo(  
        "libreoffice7.3",  
        string.Format("--invisible --convert-to pdf  {0}", filename + extension))  
    {RedirectStandardOutput = true};  
  
// 启动  
var proc = Process.Start(psi);  
if (proc == null)  
{  
    logger.Error("请检查 LibreOffice 是否成功安装.");  
    return false;  
}  
  
logger.Log(LogLevel.Info, "文件转换开始......");  
using (var sr = proc.StandardOutput)  
{  
    while (!sr.EndOfStream)  
    {        Console.WriteLine(sr.ReadLine());  
    }    if (!proc.HasExited)  
    {        proc.Kill();  
    }}  
  
logger.Log(LogLevel.Info, "文件转成完成");
```

4、文件转换成功后，存储转换后的 PDF 文件到 MongoDB，然后和原始文件进行关联，下面代码是调用了零代码平台中的接口进行处理，这里可以根据自己的业务需求自行修改 :

```
string host = EnvironmentHelper.GetEnvValue("ApiHost");  
string api = EnvironmentHelper.GetEnvValue("AssociationApi");  
if (string.IsNullOrEmpty(api))  
{  
    logger.Warn("请检查 AssociationApi 环境变量的配置");  
    return false;  
}  
if (string.IsNullOrEmpty(host))  
{  
    logger.Warn("请检查 ApiHost 环境变量的配置");  
    return false;  
}  
string result = APIHelper.RunApiGet(host, $"{api}/{fileId}/{destFileId}");
```

### 构建 Libreoffice 基础镜像

1、在 centos 服务器上 /data 目录中创建目录 liberoffice-docker-build ,将上面提到的 Libreoffice 安装包、Libreoffice 中文语言包、jdk1.8 安装包 、中文字体包拷贝到该目录中；

2、在该目录中创建 Dockerfile 文件，内容如下：

```
RUN yum update -y && \
        yum reinstall -y glibc-common && \
        yum install -y telnet net-tools && \
        yum clean all && \
        rm -rf /tmp/* rm -rf /var/cache/yum/* && \
        localedef -c -f UTF-8 -i zh_CN zh_CN.UTF-8 && \
        ln -sf /usr/share/zoneinfo/Asia/Shanghai /etc/localtime

#加入windows字体包
ADD chinese.tar.gz /usr/share/fonts/

ADD LibreOffice_7.3.5_Linux_x86-64_rpm.tar.gz /home/
ADD LibreOffice_7.3.5_Linux_x86-64_rpm_langpack_zh-CN.tar.gz /usr/

#执行安装
RUN cd /home/LibreOffice_7.3.5.2_Linux_x86-64_rpm/RPMS/ \
        && yum localinstall *.rpm -y \
        && cd /usr/LibreOffice_7.3.5.2_Linux_x86-64_rpm_langpack_zh-CN/RPMS/   \
        && yum localinstall *.rpm -y \

        #安装依赖
        && yum install ibus -y \

        #加入中文字体支持并赋权限
        && cd /usr/share/fonts/ \
        && chmod -R 755 /usr/share/fonts \
        && yum install mkfontscale -y \
        && mkfontscale \
        && yum install fontconfig -y \
        && mkfontdir \
        && fc-cache -fv \
        && mkdir /usr/local/java/ \

        #清理缓存,减少镜像大小
        && yum clean all

#安装java环境
ADD jdk-8u341-linux-x64.tar.gz /usr/local/java/
RUN ln -s /usr/local/java/jdk1.8.0_314 /usr/local/java/jdk

#配置环境变量
ENV JAVA_HOME /usr/local/java/jdk
ENV JRE_HOME ${JAVA_HOME}/jre
ENV CLASSPATH .:${JAVA_HOME}/lib:${JRE_HOME}/lib
ENV PATH ${JAVA_HOME}/bin:$PATH

#安装 dotnet core 3.1 运行环境
RUN rpm -Uvh https://packages.microsoft.com/config/centos/7/packages-microsoft-prod.rpm \
    && yum install -y  aspnetcore-runtime-3.1 \
    && yum clean all

WORKDIR /usr
EXPOSE 80
CMD /bin/bash
```

3、执行命令 `docker build -t libreofficebase:v1.0 .` 进行基础镜像的构建，构建好的基础镜像供文件预览镜像构建时使用。

### 构建文件预览镜像

1、在 centos 服务器的 /data 目录中创建目录 doc-preview-docker-build ；

2、将转换程序 OfficeToPdf 进行编译发布，将发布后的文件拷贝到目录 doc-preview-docker-build 中；

3、在该目录中创建 Dockerfile 文件，内容如下：

```
FROM libreofficebase:v1 #此处的镜像就是上面构建的 Libreoffice 基础镜像
COPY . /app
WORKDIR /app
EXPOSE 80/tcp
ENTRYPOINT ["dotnet", "OfficeToPdf.dll"]
```

4、执行命令 `docker build -t office-preview:v1.0 .` 进行预览镜像的构建。

### 运行预览容器

执行下面命令进行容器的创建：

```
docker run -d --name office-preview office-preview
```
