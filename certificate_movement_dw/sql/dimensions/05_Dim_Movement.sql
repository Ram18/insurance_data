/* =============================================================================
   Dim_Movement  — the transition dimension
   Source: dbo.MOVEMENT_MEASURE_LOOKUP
   Business key: (ADW_STATUS_KEY_START, ADW_STATUS_KEY_END)      SCD: Type 1
   Maps every start->end status transition to a movement label + flags.
   ============================================================================= */

IF OBJECT_ID(N'star.Dim_Movement', N'U') IS NULL
BEGIN
    CREATE TABLE star.Dim_Movement
    (
        MovementKey           INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Dim_Movement PRIMARY KEY,
        ADW_STATUS_KEY_START  INT            NOT NULL,   -- business key part 1
        ADW_STATUS_KEY_END    INT            NOT NULL,   -- business key part 2
        Movement              NVARCHAR(100)  NULL,       -- 'Death','Surrendered','Lapsed',...
        MovementGroup         NVARCHAR(100)  NULL,       -- 'New','Revived','Correction','Lost',...
        StartFlag             INT            NOT NULL DEFAULT 0,
        EndFlag               INT            NOT NULL DEFAULT 0,
        NewFlag               INT            NOT NULL DEFAULT 0,
        RevivedFlag           INT            NOT NULL DEFAULT 0,
        ExitFlag              INT            NOT NULL DEFAULT 0,
        LostFlag              INT            NOT NULL DEFAULT 0,
        CorrectionFlag        INT            NOT NULL DEFAULT 0,
        PremiumPayingAtStart  NVARCHAR(50)   NULL,
        CONSTRAINT UQ_Dim_Movement_BK UNIQUE (ADW_STATUS_KEY_START, ADW_STATUS_KEY_END)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM star.Dim_Movement WHERE MovementKey = -1)
BEGIN
    SET IDENTITY_INSERT star.Dim_Movement ON;
    INSERT INTO star.Dim_Movement (MovementKey, ADW_STATUS_KEY_START, ADW_STATUS_KEY_END, Movement, MovementGroup)
    VALUES (-1, -1, -1, N'Unknown', N'Unknown');
    SET IDENTITY_INSERT star.Dim_Movement OFF;
END
GO

CREATE OR ALTER VIEW star.stg_Dim_Movement AS
SELECT
    ADW_STATUS_KEY_START = m.ADW_STATUS_KEY_START,
    ADW_STATUS_KEY_END   = m.ADW_STATUS_KEY_END,
    Movement             = LTRIM(RTRIM(m.MOVEMENT)),
    MovementGroup        = LTRIM(RTRIM(m.MOVEMENT_GROUP)),
    StartFlag            = ISNULL(m.START_FLAG, 0),
    EndFlag              = ISNULL(m.END_FLAG, 0),
    NewFlag              = ISNULL(m.NEW_FLAG, 0),
    RevivedFlag          = ISNULL(m.REVIVED_FLAG, 0),
    ExitFlag             = ISNULL(m.EXIT_FLAG, 0),
    LostFlag             = ISNULL(m.LOST_FLAG, 0),
    CorrectionFlag       = ISNULL(m.CORRECTION_FLAG, 0),
    PremiumPayingAtStart = LTRIM(RTRIM(m.PREMIUM_PAYING_AT_START))
FROM dbo.MOVEMENT_MEASURE_LOOKUP AS m;
GO

MERGE star.Dim_Movement AS tgt
USING star.stg_Dim_Movement AS src
    ON  tgt.ADW_STATUS_KEY_START = src.ADW_STATUS_KEY_START
    AND tgt.ADW_STATUS_KEY_END   = src.ADW_STATUS_KEY_END
WHEN MATCHED THEN UPDATE SET
    tgt.Movement = src.Movement, tgt.MovementGroup = src.MovementGroup,
    tgt.StartFlag = src.StartFlag, tgt.EndFlag = src.EndFlag, tgt.NewFlag = src.NewFlag,
    tgt.RevivedFlag = src.RevivedFlag, tgt.ExitFlag = src.ExitFlag, tgt.LostFlag = src.LostFlag,
    tgt.CorrectionFlag = src.CorrectionFlag, tgt.PremiumPayingAtStart = src.PremiumPayingAtStart
WHEN NOT MATCHED BY TARGET THEN
    INSERT (ADW_STATUS_KEY_START, ADW_STATUS_KEY_END, Movement, MovementGroup, StartFlag, EndFlag,
            NewFlag, RevivedFlag, ExitFlag, LostFlag, CorrectionFlag, PremiumPayingAtStart)
    VALUES (src.ADW_STATUS_KEY_START, src.ADW_STATUS_KEY_END, src.Movement, src.MovementGroup, src.StartFlag,
            src.EndFlag, src.NewFlag, src.RevivedFlag, src.ExitFlag, src.LostFlag, src.CorrectionFlag,
            src.PremiumPayingAtStart);
GO
