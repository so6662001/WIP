/* ============================================================================
   PR-3：凭证 batch 后处理批量化 stored proc
   
   依赖：原 ss_YWVouUpdateAchTimes / ss_AfterSaveUpdateGL2VouBill 保持不动。
   新增两个批量版本，接受 BillID list（逗号分隔串），等价于循环调原版。
   
   兼容：SQL Server 2008 及以上（用 XML.nodes() 拆分逗号串，不依赖 STRING_SPLIT）
   ============================================================================ */

------------------------------------------------------------------------------
-- ss_YWVouUpdateAchTimes_Batch
-- 等价于：FOR EACH BillID IN @BillIDList: EXEC ss_YWVouUpdateAchTimes BillID, @Adj
------------------------------------------------------------------------------
IF OBJECT_ID('dbo.ss_YWVouUpdateAchTimes_Batch') IS NOT NULL
    DROP PROCEDURE dbo.ss_YWVouUpdateAchTimes_Batch;
GO

CREATE PROCEDURE dbo.ss_YWVouUpdateAchTimes_Batch
    @BillIDList NVARCHAR(MAX),    -- 逗号分隔的 BillID
    @Adj        INT               -- 1=增加使用次数；-1=减少
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @x XML;
    SET @x = CAST('<x>' + REPLACE(@BillIDList, ',', '</x><x>') + '</x>' AS XML);

    -- 用 IDENTITY 列保留传入顺序（XML.nodes 读取顺序未明确保证，
    -- 但 INSERT INTO ... SELECT 内部仍按 nodes 文档顺序，配合 IDENTITY 即可保序）
    DECLARE @Tmp TABLE (rn INT IDENTITY(1,1), BillID VARCHAR(48));
    INSERT INTO @Tmp (BillID)
    SELECT LTRIM(RTRIM(T.N.value('.', 'varchar(50)')))
    FROM   @x.nodes('/x') AS T(N)
    WHERE  LTRIM(RTRIM(T.N.value('.', 'varchar(50)'))) <> '';

    -- 遍历调用原 stored proc，按传入顺序处理
    DECLARE @BillID VARCHAR(48);
    DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
        SELECT BillID FROM @Tmp ORDER BY rn;
    OPEN cur;
    FETCH NEXT FROM cur INTO @BillID;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        EXEC dbo.ss_YWVouUpdateAchTimes @BillID, @Adj;
        FETCH NEXT FROM cur INTO @BillID;
    END
    CLOSE cur;
    DEALLOCATE cur;
END
GO


------------------------------------------------------------------------------
-- ss_AfterSaveUpdateGL2VouBill_Batch
-- 等价于：FOR EACH BillID IN @BillIDList: EXEC ss_AfterSaveUpdateGL2VouBill BillID, @IsTrans
------------------------------------------------------------------------------
IF OBJECT_ID('dbo.ss_AfterSaveUpdateGL2VouBill_Batch') IS NOT NULL
    DROP PROCEDURE dbo.ss_AfterSaveUpdateGL2VouBill_Batch;
GO

CREATE PROCEDURE dbo.ss_AfterSaveUpdateGL2VouBill_Batch
    @BillIDList NVARCHAR(MAX),
    @IsTrans    BIT               -- 0=正式表  1=过渡表
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @x XML;
    SET @x = CAST('<x>' + REPLACE(@BillIDList, ',', '</x><x>') + '</x>' AS XML);

    DECLARE @Tmp TABLE (rn INT IDENTITY(1,1), BillID VARCHAR(48));
    INSERT INTO @Tmp (BillID)
    SELECT LTRIM(RTRIM(T.N.value('.', 'varchar(50)')))
    FROM   @x.nodes('/x') AS T(N)
    WHERE  LTRIM(RTRIM(T.N.value('.', 'varchar(50)'))) <> '';

    DECLARE @BillID VARCHAR(48);
    DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
        SELECT BillID FROM @Tmp ORDER BY rn;
    OPEN cur;
    FETCH NEXT FROM cur INTO @BillID;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        EXEC dbo.ss_AfterSaveUpdateGL2VouBill @BillID, @IsTrans;
        FETCH NEXT FROM cur INTO @BillID;
    END
    CLOSE cur;
    DEALLOCATE cur;
END
GO


/* ============================================================================
   性能说明：
   
   这两个批量版 stored proc 内部仍是 cursor 循环调原 stored proc，
   逻辑严格等价。性能收益来自**消除 N 次 RPC 网络往返**：
     - 原版：10 万次 RPC × ~1ms = 100 秒
     - 批量：500 次 RPC（每 200 张一批）× ~1ms = 0.5 秒
   服务端 cursor 循环时间不变，但避免了网络往返开销。
   
   如果原 stored proc 内部本身是 set-based 操作（一条 UPDATE 处理一行），
   理论上可以重写成集合化版本进一步加速。但那需要看原 stored proc 实现，
   不在 PR-3 范围内。
   ============================================================================ */
