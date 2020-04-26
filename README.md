# imgToWord

### 需求说明

- 存在多个文件夹, 将其中每一个文件夹里面的图片转换到word中

- 同时将目录名的前四位放到word的第一行

- 使用多模板, 方便调整,适配每一种的特殊调整

- 图片在4-16张之间

  

### 技术说明

- 使用freemarker.jar使用word转出的(2003版)xml文件作为模板
- 由于freemarker转的word格式在手机上不能打开, 因此使用jacob调用本地的office将word另保存一份

### 其他说明

- jacob要将对应版本的jacob-1.18-x64.dll 放到jdk/jre/bin (jacob-1.18.x.zip里面包括对应dll)
- 用${xx} 来替换word中对应替换的内容
- 图片要用base64转码