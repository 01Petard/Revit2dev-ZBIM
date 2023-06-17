# ZBIMUtils

## 简介
Revit二次开发项目常用公共方法库



## 要求

- .NET Framework == **4.5.2**
- RevitAPI == **18.0.0.0**
- RevitAPIUI == **18.0.0.0**



## 使用说明

**工程师：**

1. 在Deployment文件夹中获取ZBIMUtils.dll公共方法库；
2. 将ZBIMUtils.dll和待测强条对应的dll文件（如ZBIM-H00001.dll）放在同一目录下；
3. 在 Revit2018 -> 附加模块 -> 外部工具 中，加载待测强条对应的dll，然后进行测试。

**开发者：**

1. 在VS项目-引用中右键 -> 添加引用；

   ![](./img/1-%E6%B7%BB%E5%8A%A0%E5%BC%95%E7%94%A8.jpg)

2. 在引用管理器-浏览中选择路径./Deployment，选中并勾选ZBIMUtils.dll；

   ![](./img/2-%E5%BC%95%E7%94%A8%E7%AE%A1%E7%90%86%E5%99%A8.jpg)

3. 在名称空间中添加ZBIMUtils；

   ```C#
   using ZBIMUtils;
   ```

4. 根据**函数清单（见下方）**，选择需要的公共方法进行调用；

   ```C#
   double height_m = Utils.FootToMeter(height_ft)
   ```

5. 注：**ZBIMUtils.dll和ZBIMUtils.xml需要放在同一目录下**。




## 常见问题

### 问题1：未能加载文件或程序集ZBIMUtils.dll

解决方案：选中ZBIMUtils.dll，右键->属性，若该dll被锁点，点击解除锁定，如下图所示，然后重新运行。

![](./img/dll%E8%B0%83%E7%94%A8%E9%97%AE%E9%A2%98.jpg)



## 函数清单

### 单位换算相关

| 函数名                  | 简介                          | 备注 |
| ----------------------- | ----------------------------- | ---- |
| FootToMeter             | 单位换算：英尺(ft) => 米(m)   |      |
| MeterToFoot             | 单位换算：米(m) => 英尺(ft)   |      |
| SquareFootToSquareMeter | 平方英尺(ft^2) => 平方米(m^2) |      |
| SquareMeterToSquareFoot | 平方米(m^2) => 平方英尺(ft^2) |      |
| CubicFootToCubicMeter   | 立方英尺(ft^3) => 立方米(m^3) |      |
| CubicMeterToCubicFoot   | 立方米(m^3) => 立方英尺(ft^3) |      |

<u>注：公共方法库中提供的方法，默认单位均采用**英制**，如有需要，请根据上述方法进行换算。</u>



### 字符串操作相关

| 函数名           | 简介                 | 备注 |
| ---------------- | -------------------- | ---- |
| StartSubString   | (截头)截取字符串     |      |
| SubBetweenString | (截中)截取字符串     |      |
| LastSubEndString | 反向截末——截取字符串 |      |



### 建筑高度、标高相关

| 函数名                   | 简介                           | 备注 |
| ------------------------ | ------------------------------ | ---- |
| GetBuildingBoundingBoxUV | 获取建筑的楼号和对应的边界框   |      |
| GetBuildingHeight        | 获取建筑的楼号和对应的高度     |      |
| GetBasementHeight        | 获取地下室的高度               |      |
| GetNumFromLevelName      | 通过标高命名获取其相关联的楼号 |      |
| GetBuildingLevels        | 获取建筑的楼号和对应的标高     |      |
| GetOutdoorTerrace        | 获取室外地坪                   |      |
| InBuilding               | 判断元素是否在建筑内           |      |



### 门窗相关

| 函数名                    | 简介                         | 备注                 |
| ------------------------- | ---------------------------- | -------------------- |
| GetWindowsInRoom          | 获取房间内所有的窗户         |                      |
| GetDoorsInRoom            | 获取房间内所有的门           |                      |
| IsRequiredDoorsFireRating | 判断门的防火等级是否符合要求 |                      |
| IsExternalDoor            | 判断门的一侧是否朝向室外     |                      |
| IsExternalWindow          | 判断窗户是否为外窗           |                      |
| IsOpenableWindow          | 判断窗户是否为可开启窗       |                      |
| IsInstFromRoom            | 判断门窗是否属于某个房间     | 利用空间几何关系判断 |
| GetWindowOpenableArea     | 计算窗户的可开启面积         |                      |



### 房间相关

| 函数名                  | 简介                         | 备注 |
| ----------------------- | ---------------------------- | ---- |
| GetRoomIntersectElement | 获取房间内的所有元素         |      |
| IsEnclosedBalcony       | 判断阳台是否为封闭式         |      |
| GetRoomNetHeight        | 计算房间的净高               |      |
| GetSmokeExhaustArea     | 计算房间内的有效排烟面积之和 |      |
| GetSmokeControlArea     | 计算房间内的防烟面积之和     |      |
| GetRoomNeighbours       | 查找相邻房间                 |      |
| GetPartitionWalls       | 获取房间内的隔墙             |      |

#### 如何提高获取房间内部元素的速度：

```c#
Dictionary<string, BoundingBoxUV> buildingBbox = Utils.
    GetBuildingBoundingBoxUV(
    doc, 
    uiDoc, 
    expand: 1.2, 
    matchAll: false);

Dictionary<string, List<Level>> buildingLev = Utils.
    GetBuildingLevels(
    buildingBbox.Keys.ToList(), 
    doc);

Dictionary<string, <ElementId, List<Room>>> GetRoomDict(
    List<Room> targetRooms,
    Dictionary<string, BoundingBoxUV> buildingBbox,
    Dictionary<string, List<Level>> buildingLev)
{
	// roomDict通过楼号<string>和楼层ID<ElementId>检索指定位置的房间
	var roomDict = new Dictionary<string, Dictionary<ElementId, List<Room>>>();
    
    /* 
     * 根据buildingBbox中的楼栋范围，和buildingLev中的楼层信息，
     * 将房间按照楼号和楼层进行分组。
     * ......
     * 最后得到完整的roomDict。
     */
    
    return roomDict;
}

Dictionary<string, <ElementId, List<ElementId>>> GetElemIdDict(
    List<ElementId> targetElemIds,
    Dictionary<string, BoundingBoxUV> buildingBbox,
    Dictionary<string, List<Level>> buildingLev)
{
	// elemIdDict通过楼号<string>和楼层ID<ElementId>检索指定位置的元素ID
	var elemIdDict 
        = new Dictionary<string, Dictionary<ElementId, List<ElementId>>>();
    
    /* 
     * 同上，
     * ......
     * 最后得到完整的elemIdDict。
     */
    
    return elemIdDict;
}

var roomDict = GetRoomDict(...);
var elemIdDict = GetElemIdDict(...);

// 按照楼号-楼层的结构进行遍历
foreach(var num in roomDict.Keys)
{
    foreach(var levId in roomDict[num].Keys)
    {
        List<ElementId> searchRange = elemIdDict[num][levId];
        foreach(var room in roomDcit[num][levId])
        {
            // 通过缩小搜索范围来提高获取的效率
            FilteredElementCollector interElems = 
                GetRoomIntersectElement(
                room, 
                doc, 
                scaleFactor:1,
                useBoundingBox:false,
                searchRange:searchRange);           
            /* 
             * 其他处理
             * ......
             */            
        }
    }
}
```



### 管道相关

| 函数名            | 简介             | 备注 |
| ----------------- | ---------------- | ---- |
| GetPipesBySysName | 按系统名过滤管道 |      |
| IsVerticalPipe    | 判断是否为立管   |      |



### 视图切换相关

| 函数名           | 简介           | 备注 |
| ---------------- | -------------- | ---- |
| SwitchTo2D       | 切换到二维视图 |      |
| SwitchTo3D       | 切换到三维视图 |      |
| CloseCurrentView | 关闭当前视图   |      |

#### 视图转换使用范例：

```C#
View preView = null;
bool switched = false;
Utils.SwitchTo2D(doc, uiDoc, ref preView, ref switched, viewName:"1F");

/* 
 * Write your code here. 
 * ......
 */

// 关闭当前视图，即切换回原视图
Utils.CloseCurrentView(uiDoc);
```



### 防火分区相关

| 函数名             | 简介                                             | 备注 |
| ------------------ | ------------------------------------------------ | ---- |
| GetFireAreas       | 获取所有的防火分区                               |      |
| GetFireAreaAtPoint | 获取某点所在的防火分区（与GetFireAreas结合使用） |      |



### 其他

| 函数名              | 简介                  | 备注 |
| ------------------- | --------------------- | ---- |
| PrintLog            | 打印log文件到指定目录 |      |
| ScaleBoundingBoxXYZ | 缩放三维边界框        |      |
| GetOtherOpenings    | 获取除门窗外的洞口    |      |

#### 打印log使用范例：

```C#
string log = "";

/* 
 * Write the information that needs to be presented to the "log".
 * ......
 */

Utils.PrintLog(log, "H00xxx", doc);
```



## 更新日志

v230615:

1. 删除了GetBuildingBoundingBoxUV中活动视图切换；

v230614:

1. 改变了GetFireAreaAtPoint的底层逻辑；
2. 修复了GetWindowsInRoom无法获取部分幕墙窗户的bug；

v230613:

1. 修复了建筑高度计算的bug；

v230612:

1. 增加了GetOtherOpenings；
3. 增加了GetFireAreas、GetFireAreaAtPoint；

v230601:

1. 修复了房间以幕墙为边界墙的情况下，GetWindowsInRoom和GetDoorsInRoom无法获取到窗门的bug；
2. 考虑到部分门窗即便两侧均有房间，toRoom和fromRoom也有可能为null，本次更新修改了GetWindowsInRoom和GetDoorsInRoom中的门窗与房间从属关系的判别方法，使其在上述情况下也能正常进行判断，新的判别方法封装在IsInstFromRoom中；
3. 增加了GetWindowOpenableArea，可计算窗户的可开启面积；

v230511:

1. 修改GetPipesBySysName的获取规则，支持同时获取OST_PipeCurves、OST_PipeAccessory、OST_PipeFitting三种类型的管道；

v230506:

1. 增加GetPartitionWalls；

v230417:

1. GetRoomIntersectsElement增加缩放因子scaleFactor；
1. 增加立管判断IsVerticalPipe；
1. 增加获取室外地坪GetOutdoorTerrace；
1. 修复了查找相邻房间GetRoomNeighbours中处理特殊房间会报错中断的bug。

v230408:

1. GetBuildingHeight修改：机房屋面和屋面使用结构标高、女儿墙使用建筑标高；增加了函数注释，列举了返回值为-1的异常情况；
2. 增加了InBuilding，用于判断元素是否在指定建筑内；
3. 更改了SwitchTo2D，现已支持切换到指定名称的二维视图；增加了SwitchTo3D、CloseCurrentView；删除了SwitchBack；
4. 增加了GetPipesBySysName，用于通过管道系统简称过滤管道；
5. 增加了字符串操作StartSubString、SubBetweenString、LastSubEndString 。

v230404 (上线gitee):

1. 增加了IsEnclosedBalcony、IsExternalDoor、GetNumFromLevelName、GetBuildingLevels、GetRoomNeighbours、SquareFootToSquarMeter、SquareMeterToSquarFoot、CubicFootToCubicMeter、CubicMeterToCubicFoot。

v230403: 

1. 初试版本。

v230403_2: 

1. 增加了视图转换SwitchTo2D、SwitchBack。

v230403_1: 

1. 增加函数注释 (xml需要和.dll需要放在同一目录下)；

2. GetRoomIntersectsElement增加优化项searchRange。
