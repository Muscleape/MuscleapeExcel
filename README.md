# *MuscleapeExcel* is a excel tool written in Java base on [alibaba easyexcel](https://github.com/alibaba/easyexcel)

## 为啥建这个仓库

> 1、在项目中使用了easyexcel，并修改了部分源码来满足业务需求；
>
> 2、作为自己学习源码的一个基础；
>
> 3、参考网上针对其他队easyexcel的封装打造自己的一个项目，方便组内小伙伴们使用；


## 变更记录

- （2019-04-25）针对无模板导出，实现键值(枚举值、动态值)转换导出，例如：数据0/1导出到Excel中为男/女；
- （2019-04-26）修改项目名称及包名等信息；
- （2019-04-27）使用Lombok并全局修改，移除部分过时方法，更改开发者信息；
- （2019-04-28）精简（部分与POI重名的类重新命名，部分工具类改用spring和common-lang3的）；
- （2019-05-01）添加一个工具类【WebWriteUtil.class】，方便WEB项目中浏览器下载生成的Excel文件，该工具类可以直接放到web项目中使用；