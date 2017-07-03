--付款申请单
--直接在BOS中调用

--注册销售订单获取应收账款插件
INSERT INTO t_ThirdPartyComponent VALUES(0,81,'SM.Industry_XSDD','','获取应收账款')
--注册销售订单中间层插件
INSERT INTO t_ThirdPartyComponent VALUES(2,200001,2001,'SM.Industry_MiddlewareEvent_XSDD','','单价未零时不允许审核')
--付款申请单可在BOS注册
--更新插件时间戳
UPDATE t_DataFlowTimeStamp SET FName=FName

--序时簿插入按钮步骤
--1. 在表 t_MenuToolBar 插入工具栏按钮
--2. 在表 ICListTemplate 查询对应序时簿的 FMenuID
--3. 在表 t_BandToolMapping 中插入按钮，FID为2的FMenuID，
--      FBandID为你要插入的位置（菜单，工具栏），
--      FIndex为要插入的顺序，小左，大右，
--      FComName为插件名称，前面要加" | "
--查询语句
SELECT
	T.FName,
	M.FBandID,
	M.FIndex,
	M.FToolID,
	M.FComName,
	B.FToolCaption
FROM ICListTemplate AS T
INNER JOIN t_BandToolMapping AS M ON T.FMenuID=M.FID
INNER JOIN t_MenuToolBar AS B ON B.FToolID=M.FToolID
WHERE T.FName LIKE '%销售订单%'
ORDER BY T.FID,M.FBandID,M.FIndex,M.FToolID