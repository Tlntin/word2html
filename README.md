## 为啥要做这个软件

- 部分编辑器排版不是很方便，但对html代码支持程度较好。

## word不是直接能装html了嘛？你这个有啥不一样

- word自带转html代码不是很好，内嵌了很多字体和颜色格式，而这些一般第三方编辑器并不支持。

- 我这个直接就是比较简单的html代码，比如`h1`、`em`、`p`之类的简单代码，理论上支持大部分第三方html格式。

## 软件特色

- 基于PyQt5打包，所以软件比较大，30多M。
- 基于`mammoth`模块开发，这个提供了doc,docx转html的核心

## 打包方式

```shell
pyinstaller -i png.ico -F -w main.py
```

## 需要的模块

```shell
pip install PyQt5
pip install pyperclip
pip install mammoth
```



