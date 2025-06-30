/**********************************************************************/
/* UpgradeMembershipProvider.sql                                      */
/*                                                                    */
/* Upgrades the tables, triggers and stored procedures necessary for  */
/* supporting the aspnet feature of ASP.Net                           */
/**********************************************************************/

/*************************************************************/
/* aspnet Updates											 */
/*************************************************************/

SET NUMERIC_ROUNDABORT OFF
GO
SET ANSI_PADDING, ANSI_WARNINGS, CONCAT_NULL_YIELDS_NULL, ARITHABORT, QUOTED_IDENTIFIER, ANSI_NULLS ON
GO
SET XACT_ABORT ON
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO

/*************************************************************/
/* Table Updates											 */
/*************************************************************/
--Remove all the Constraints, we don't know their exact names but we do know how the name starts
DECLARE @name varchar(64)

--Remove Foreign Keys
SET @name = (SELECT name from sysobjects WHERE name Like 'FK__aspnet_Me__UserI%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Membership] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'FK__aspnet_Pr__UserI%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Profile] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'FK__aspnet_Ro__Appli%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Roles] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'FK__aspnet_Us__RoleI%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_UsersInRoles] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'FK__aspnet_Us__UserI%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_UsersInRoles] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'FK__aspnet_Us__Appli%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Users] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'FK__aspnet_Ro__Appli%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Roles] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'FK__aspnet_Us__Appli%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Users] DROP CONSTRAINT ' + @name)

--Remove Default Constraints
SET @name = (SELECT name from sysobjects WHERE name Like 'DF__aspnet_Me__Passw%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Membership] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'DF__aspnet_Pr__LastU%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Profile] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'DF__aspnet_Ap__Appli%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Applications] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'DF__aspnet_Ro__RoleI%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Roles] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'DF__aspnet_Us__UserI%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Users] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'DF__aspnet_Us__Mobil%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Users] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'DF__aspnet_Us__IsAno%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Users] DROP CONSTRAINT ' + @name)

--Remove Unique Constraints
SET @name = (SELECT TOP 1 name from sysobjects WHERE name Like 'UQ__aspnet_Applicati%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Applications] DROP CONSTRAINT ' + @name)
SET @name = (SELECT TOP 1 name from sysobjects WHERE name Like 'UQ__aspnet_Applicati%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Applications] DROP CONSTRAINT ' + @name)

--Remove Primary Keys
SET @name = (SELECT TOP 1 name from sysobjects WHERE name Like 'PK__aspnet_Users%' ORDER BY name)
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Users] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'PK__aspnet_Membershi%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Membership] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'PK__aspnet_Profile%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Profile] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'PK__aspnet_Roles%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Roles] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'PK__aspnet_UsersInRo%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_UsersInRoles] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'PK__aspnet_Applicati%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_Applications] DROP CONSTRAINT ' + @name)
SET @name = (SELECT name from sysobjects WHERE name Like 'PK__aspnet_SchemaVer%')
IF Not @name Is Null
	EXEC ('ALTER TABLE [dbo].[aspnet_SchemaVersions] DROP CONSTRAINT ' + @name)

--End removal of Constraints
GO

ALTER TABLE [dbo].[aspnet_Profile] ALTER COLUMN [LastUpdatedDate] [datetime] NOT NULL
GO

DROP INDEX [dbo].[aspnet_UsersInRoles].[aspnet_UsersInRoles_index2]
GO
DROP INDEX [dbo].[aspnet_Membership].[aspnet_Membership_index]
GO
DROP INDEX [dbo].[aspnet_Roles].[aspnet_Roles_index1]
GO
DROP INDEX [dbo].[aspnet_Users].[aspnet_Users_Index]
GO
DROP INDEX [dbo].[aspnet_UsersInRoles].[aspnet_UsersInRoles_index1]
GO
DROP VIEW [dbo].[vw_aspnet_SchemaVersions]
GO

/*************************************************************/
/* Update Membership Table retaining existing Data			 */
/*************************************************************/
--Create Temp Table
CREATE TABLE [dbo].[tmp_rg_xx_aspnet_Membership]
(
	[ApplicationId] [uniqueidentifier] NOT NULL,
	[UserId] [uniqueidentifier] NOT NULL,
	[Password] [nvarchar] (128) NOT NULL,
	[PasswordFormat] [int] NOT NULL CONSTRAINT [DF__aspnet_Membership_PasswordFormat] DEFAULT (0),
	[PasswordSalt] [nvarchar] (128) NOT NULL,
	[MobilePIN] [nvarchar] (16) NULL,
	[Email] [nvarchar] (256) NULL,
	[LoweredEmail] [nvarchar] (256) NULL,
	[PasswordQuestion] [nvarchar] (256) NULL,
	[PasswordAnswer] [nvarchar] (128) NULL,
	[IsApproved] [bit] NOT NULL,
	[IsLockedOut] [bit] NOT NULL,
	[CreateDate] [datetime] NOT NULL,
	[LastLoginDate] [datetime] NOT NULL,
	[LastPasswordChangedDate] [datetime] NOT NULL,
	[LastLockoutDate] [datetime] NOT NULL,
	[FailedPasswordAttemptCount] [int] NOT NULL,
	[FailedPasswordAttemptWindowStart] [datetime] NOT NULL,
	[FailedPasswordAnswerAttemptCount] [int] NOT NULL,
	[FailedPasswordAnswerAttemptWindowStart] [datetime] NOT NULL,
	[Comment] [ntext] NULL
)
GO

--Insert Existing Data into Temp Table
INSERT INTO [dbo].[tmp_rg_xx_aspnet_Membership]
	([ApplicationId],
	[UserId], 
	[Password], 
	[PasswordFormat], 
	[PasswordSalt], 
	[MobilePIN], 
	[Email], 
	[LoweredEmail], 
	[PasswordQuestion], 
	[PasswordAnswer], 
	[IsApproved], 
	[IsLockedOut], 
	[CreateDate], 
	[LastLoginDate], 
	[LastPasswordChangedDate], 
	[LastLockoutDate], 
	[FailedPasswordAttemptCount], 
	[FailedPasswordAttemptWindowStart], 
	[FailedPasswordAnswerAttemptCount], 
	[FailedPasswordAnswerAttemptWindowStart], 
	[Comment]) 
SELECT 	[ApplicationId],
	m.[UserId], 
	[Password], 
	[PasswordFormat], 
	[PasswordSalt], 
	[MobilePIN], 
	[Email], 
	[LoweredEmail], 
	[PasswordQuestion], 
	[PasswordAnswer], 
	[IsApproved], 
	[IsLockedOut], 
	[CreateDate], 
	[LastLoginDate], 
	[LastPasswordChangedDate], 
	[LastLockoutDate], 
	[FailedPasswordAttemptCount], 
	[FailedPasswordAttemptWindowStart], 
	[FailedPasswordAnswerAttemptCount], 
	[FailedPasswordAnswerAttemptWindowStart], 
	[Comment] 
FROM    [dbo].[aspnet_Users] u INNER JOIN
        [dbo].[aspnet_Membership] m ON u.UserId = m.UserId
GO

DROP TABLE [dbo].[aspnet_Membership]
GO

sp_rename N'[dbo].[tmp_rg_xx_aspnet_Membership]', N'aspnet_Membership'
GO

CREATE CLUSTERED INDEX [aspnet_Membership_index] ON [dbo].[aspnet_Membership] ([ApplicationId], [LoweredEmail])
ALTER TABLE [dbo].[aspnet_Membership] ADD CONSTRAINT [PK__aspnet_Membership] PRIMARY KEY NONCLUSTERED  ([UserId])
GO

/*************************************************************/
/*************************************************************/

CREATE PROCEDURE [dbo].aspnet_GetUtcDate
    @TimeZoneAdjustment INT,
    @DateNow            DATETIME OUTPUT
AS
BEGIN
    SELECT @DateNow = GETUTCDATE()
END
GO

ALTER PROCEDURE [dbo].aspnet_CheckSchemaVersion
    @Feature                   NVARCHAR(128),
    @CompatibleSchemaVersion   NVARCHAR(128)
AS
BEGIN
    IF (EXISTS( SELECT  *
                FROM    dbo.aspnet_SchemaVersions
                WHERE   Feature = LOWER( @Feature ) AND
                        CompatibleSchemaVersion = @CompatibleSchemaVersion ))
        RETURN 0

    RETURN 1
END

GO

CREATE PROCEDURE [dbo].aspnet_UnRegisterSchemaVersion
    @Feature                   NVARCHAR(128),
    @CompatibleSchemaVersion   NVARCHAR(128)
AS
BEGIN
    DELETE FROM dbo.aspnet_SchemaVersions
        WHERE   Feature = LOWER(@Feature) AND @CompatibleSchemaVersion = CompatibleSchemaVersion
END

GO

ALTER PROCEDURE dbo.aspnet_Profile_DeleteInactiveProfiles
    @ApplicationName        NVARCHAR(256),
    @ProfileAuthOptions     INT,
    @InactiveSinceDate      DATETIME,
    @TimeZoneAdjustment     INT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
    BEGIN
        SELECT  0
        RETURN
    END

    IF (@InactiveSinceDate > CONVERT(DATETIME, '17540101', 112) AND  @InactiveSinceDate < CONVERT(DATETIME, '99980101', 112))
        SELECT @InactiveSinceDate = DATEADD(n, -@TimeZoneAdjustment, @InactiveSinceDate)

    DELETE
    FROM    dbo.aspnet_Profile
    WHERE   UserId IN
            (   SELECT  UserId
                FROM    dbo.aspnet_Users u
                WHERE   ApplicationId = @ApplicationId
                        AND (LastActivityDate <= @InactiveSinceDate)
                        AND (
                                (@ProfileAuthOptions = 2)
                             OR (@ProfileAuthOptions = 0 AND IsAnonymous = 1)
                             OR (@ProfileAuthOptions = 1 AND IsAnonymous = 0)
                            )
            )

    SELECT  @@ROWCOUNT
END

GO

ALTER PROCEDURE [dbo].aspnet_Users_DeleteUser
    @ApplicationName  NVARCHAR(256),
    @UserName         NVARCHAR(256),
    @TablesToDeleteFrom INT,
    @NumTablesDeletedFrom INT OUTPUT
AS
BEGIN
    DECLARE @UserId               UNIQUEIDENTIFIER
    SELECT  @UserId               = NULL
    SELECT  @NumTablesDeletedFrom = 0

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
	    BEGIN TRANSACTION
	    SET @TranStarted = 1
    END
    ELSE
	SET @TranStarted = 0

    DECLARE @ErrorCode   INT
    DECLARE @RowCount    INT

    SET @ErrorCode = 0
    SET @RowCount  = 0

    SELECT  @UserId = u.UserId
    FROM    dbo.aspnet_Users u, dbo.aspnet_Applications a
    WHERE   u.LoweredUserName       = LOWER(@UserName)
        AND u.ApplicationId         = a.ApplicationId
        AND LOWER(@ApplicationName) = a.LoweredApplicationName

    IF (@UserId IS NULL)
    BEGIN
        GOTO Cleanup
    END

    -- Delete from Membership table if (@TablesToDeleteFrom & 1) is set
    IF ((@TablesToDeleteFrom & 1) <> 0 AND
        (EXISTS (SELECT name FROM sysobjects WHERE (name = N'aspnet_Membership') AND (type = 'U'))))
    BEGIN
        DELETE FROM dbo.aspnet_Membership WHERE @UserId = UserId

        SELECT @ErrorCode = @@ERROR,
               @RowCount = @@ROWCOUNT

        IF( @ErrorCode <> 0 )
            GOTO Cleanup

        IF (@RowCount <> 0)
            SELECT  @NumTablesDeletedFrom = @NumTablesDeletedFrom + 1
    END

    -- Delete from aspnet_UsersInRoles table if (@TablesToDeleteFrom & 2) is set
    IF ((@TablesToDeleteFrom & 2) <> 0  AND
        (EXISTS (SELECT name FROM sysobjects WHERE (name = N'aspnet_UsersInRoles') AND (type = 'U'))) )
    BEGIN
        DELETE FROM dbo.aspnet_UsersInRoles WHERE @UserId = UserId

        SELECT @ErrorCode = @@ERROR,
                @RowCount = @@ROWCOUNT

        IF( @ErrorCode <> 0 )
            GOTO Cleanup

        IF (@RowCount <> 0)
            SELECT  @NumTablesDeletedFrom = @NumTablesDeletedFrom + 1
    END

    -- Delete from aspnet_Profile table if (@TablesToDeleteFrom & 4) is set
    IF ((@TablesToDeleteFrom & 4) <> 0  AND
        (EXISTS (SELECT name FROM sysobjects WHERE (name = N'aspnet_Profile') AND (type = 'U'))) )
    BEGIN
        DELETE FROM dbo.aspnet_Profile WHERE @UserId = UserId

        SELECT @ErrorCode = @@ERROR,
                @RowCount = @@ROWCOUNT

        IF( @ErrorCode <> 0 )
            GOTO Cleanup

        IF (@RowCount <> 0)
            SELECT  @NumTablesDeletedFrom = @NumTablesDeletedFrom + 1
    END

    -- Delete from aspnet_PersonalizationPerUser table if (@TablesToDeleteFrom & 8) is set
    IF ((@TablesToDeleteFrom & 8) <> 0  AND
        (EXISTS (SELECT name FROM sysobjects WHERE (name = N'aspnet_PersonalizationPerUser') AND (type = 'U'))) )
    BEGIN
        DELETE FROM dbo.aspnet_PersonalizationPerUser WHERE @UserId = UserId

        SELECT @ErrorCode = @@ERROR,
                @RowCount = @@ROWCOUNT

        IF( @ErrorCode <> 0 )
            GOTO Cleanup

        IF (@RowCount <> 0)
            SELECT  @NumTablesDeletedFrom = @NumTablesDeletedFrom + 1
    END

    -- Delete from aspnet_Users table if (@TablesToDeleteFrom & 1,2,4 & 8) are all set
    IF ((@TablesToDeleteFrom & 1) <> 0 AND
        (@TablesToDeleteFrom & 2) <> 0 AND
        (@TablesToDeleteFrom & 4) <> 0 AND
        (@TablesToDeleteFrom & 8) <> 0 AND
        (EXISTS (SELECT UserId FROM dbo.aspnet_Users WHERE @UserId = UserId)))
    BEGIN
        DELETE FROM dbo.aspnet_Users WHERE @UserId = UserId

        SELECT @ErrorCode = @@ERROR,
                @RowCount = @@ROWCOUNT

        IF( @ErrorCode <> 0 )
            GOTO Cleanup

        IF (@RowCount <> 0)
            SELECT  @NumTablesDeletedFrom = @NumTablesDeletedFrom + 1
    END

    IF( @TranStarted = 1 )
    BEGIN
	    SET @TranStarted = 0
	    COMMIT TRANSACTION
    END

    RETURN 0

Cleanup:
    SET @NumTablesDeletedFrom = 0

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
	    ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode

END

GO

ALTER PROCEDURE [dbo].aspnet_Applications_CreateApplication
    @ApplicationName      NVARCHAR(256),
    @ApplicationId        UNIQUEIDENTIFIER OUTPUT
AS
BEGIN
    SELECT  @ApplicationId = ApplicationId FROM dbo.aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName

    IF(@ApplicationId IS NULL)
    BEGIN
        DECLARE @TranStarted   BIT
        SET @TranStarted = 0
        
        IF( @@TRANCOUNT = 0 )
        BEGIN
	        BEGIN TRANSACTION
	        SET @TranStarted = 1
        END
        ELSE
    	    SET @TranStarted = 0
        
        SELECT  @ApplicationId = ApplicationId 
        FROM dbo.aspnet_Applications WITH (UPDLOCK, HOLDLOCK) 
        WHERE LOWER(@ApplicationName) = LoweredApplicationName

        IF(@ApplicationId IS NULL)
        BEGIN
            SELECT  @ApplicationId = NEWID()
            INSERT  dbo.aspnet_Applications (ApplicationId, ApplicationName, LoweredApplicationName)
            VALUES  (@ApplicationId, @ApplicationName, LOWER(@ApplicationName))
        END
        
        
        IF( @TranStarted = 1 )
        BEGIN
            IF(@@ERROR = 0)
            BEGIN
	        SET @TranStarted = 0
	        COMMIT TRANSACTION
            END
            ELSE
            BEGIN 
                SET @TranStarted = 0
                ROLLBACK TRANSACTION
            END
        END
    END
END

GO

  ALTER VIEW [dbo].[vw_aspnet_MembershipUsers]
  AS SELECT [dbo].[aspnet_Membership].[UserId],
            [dbo].[aspnet_Membership].[PasswordFormat],
            [dbo].[aspnet_Membership].[MobilePIN],
            [dbo].[aspnet_Membership].[Email],
            [dbo].[aspnet_Membership].[LoweredEmail],
            [dbo].[aspnet_Membership].[PasswordQuestion],
            [dbo].[aspnet_Membership].[PasswordAnswer],
            [dbo].[aspnet_Membership].[IsApproved],
            [dbo].[aspnet_Membership].[IsLockedOut],
            [dbo].[aspnet_Membership].[CreateDate],
            [dbo].[aspnet_Membership].[LastLoginDate],
            [dbo].[aspnet_Membership].[LastPasswordChangedDate],
            [dbo].[aspnet_Membership].[LastLockoutDate],
            [dbo].[aspnet_Membership].[FailedPasswordAttemptCount],
            [dbo].[aspnet_Membership].[FailedPasswordAttemptWindowStart],
            [dbo].[aspnet_Membership].[FailedPasswordAnswerAttemptCount],
            [dbo].[aspnet_Membership].[FailedPasswordAnswerAttemptWindowStart],
            [dbo].[aspnet_Membership].[Comment],
            [dbo].[aspnet_Users].[ApplicationId],
            [dbo].[aspnet_Users].[UserName],
            [dbo].[aspnet_Users].[MobileAlias],
            [dbo].[aspnet_Users].[IsAnonymous],
            [dbo].[aspnet_Users].[LastActivityDate]
  FROM [dbo].[aspnet_Membership] INNER JOIN [dbo].[aspnet_Users]
      ON [dbo].[aspnet_Membership].[UserId] = [dbo].[aspnet_Users].[UserId]
  
GO
ALTER PROCEDURE dbo.aspnet_Membership_UpdateUserInfo
    @ApplicationName                NVARCHAR(256),
    @UserName                       NVARCHAR(256),
    @IsPasswordCorrect              BIT,
    @UpdateLastLoginActivityDate    BIT,
    @MaxInvalidPasswordAttempts     INT,
    @PasswordAttemptWindow          INT,
    @TimeZoneAdjustment             INT
AS
BEGIN
    DECLARE @UserId                                 UNIQUEIDENTIFIER
    DECLARE @IsApproved                             BIT
    DECLARE @IsLockedOut                            BIT
    DECLARE @LastLockoutDate                        DATETIME
    DECLARE @FailedPasswordAttemptCount             INT
    DECLARE @FailedPasswordAttemptWindowStart       DATETIME
    DECLARE @FailedPasswordAnswerAttemptCount       INT
    DECLARE @FailedPasswordAnswerAttemptWindowStart DATETIME

    DECLARE @ErrorCode     INT
    SET @ErrorCode = 0

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
	    BEGIN TRANSACTION
	    SET @TranStarted = 1
    END
    ELSE
    	SET @TranStarted = 0

    DECLARE @DateTimeNowUTC DATETIME
    EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT
    SELECT  @UserId = u.UserId,
            @IsApproved = m.IsApproved,
            @IsLockedOut = m.IsLockedOut,
            @LastLockoutDate = m.LastLockoutDate,
            @FailedPasswordAttemptCount = m.FailedPasswordAttemptCount,
            @FailedPasswordAttemptWindowStart = m.FailedPasswordAttemptWindowStart,
            @FailedPasswordAnswerAttemptCount = m.FailedPasswordAnswerAttemptCount,
            @FailedPasswordAnswerAttemptWindowStart = m.FailedPasswordAnswerAttemptWindowStart
    FROM    dbo.aspnet_Applications a, dbo.aspnet_Users u, dbo.aspnet_Membership m WITH ( UPDLOCK )
    WHERE   LOWER(@ApplicationName) = a.LoweredApplicationName AND
            u.ApplicationId = a.ApplicationId    AND
            u.UserId = m.UserId AND
            LOWER(@UserName) = u.LoweredUserName

    IF ( @@rowcount = 0 )
    BEGIN
        SET @ErrorCode = 1
        GOTO Cleanup
    END

    IF( @IsLockedOut = 1 )
    BEGIN
        GOTO Cleanup
    END

    IF( @IsPasswordCorrect = 0 )
    BEGIN
        IF( @DateTimeNowUTC > DATEADD( minute, @PasswordAttemptWindow, @FailedPasswordAttemptWindowStart ) )
        BEGIN
            SET @FailedPasswordAttemptWindowStart = @DateTimeNowUTC
            SET @FailedPasswordAttemptCount = 1
        END
        ELSE
        BEGIN
            SET @FailedPasswordAttemptWindowStart = @DateTimeNowUTC
            SET @FailedPasswordAttemptCount = @FailedPasswordAttemptCount + 1
        END

        BEGIN
            IF( @FailedPasswordAttemptCount >= @MaxInvalidPasswordAttempts )
            BEGIN
                SET @IsLockedOut = 1
                SET @LastLockoutDate = @DateTimeNowUTC
            END
        END
    END
    ELSE
    BEGIN
        IF( @UpdateLastLoginActivityDate = 1 )
        BEGIN
            UPDATE  dbo.aspnet_Membership
            SET     LastLoginDate = @DateTimeNowUTC
            WHERE   UserId = @UserId

            IF( @@ERROR <> 0 )
            BEGIN
                SET @ErrorCode = -1
                GOTO Cleanup
            END

            UPDATE  dbo.aspnet_Users
            SET     LastActivityDate = @DateTimeNowUTC
            WHERE   @UserId = UserId

            IF( @@ERROR <> 0 )
            BEGIN
                SET @ErrorCode = -1
                GOTO Cleanup
            END
        END

        IF( @FailedPasswordAttemptCount > 0 OR @FailedPasswordAnswerAttemptCount > 0 )
        BEGIN
            SET @FailedPasswordAttemptCount = 0
            SET @FailedPasswordAttemptWindowStart = CONVERT( DATETIME, '17540101', 112 )
            SET @FailedPasswordAnswerAttemptCount = 0
            SET @FailedPasswordAnswerAttemptWindowStart = CONVERT( DATETIME, '17540101', 112 )
            SET @LastLockoutDate = CONVERT( DATETIME, '17540101', 112 )
        END
    END

    UPDATE dbo.aspnet_Membership
    SET IsLockedOut = @IsLockedOut, LastLockoutDate = @LastLockoutDate,
        FailedPasswordAttemptCount = @FailedPasswordAttemptCount,
        FailedPasswordAttemptWindowStart = @FailedPasswordAttemptWindowStart,
        FailedPasswordAnswerAttemptCount = @FailedPasswordAnswerAttemptCount,
        FailedPasswordAnswerAttemptWindowStart = @FailedPasswordAnswerAttemptWindowStart
    WHERE @UserId = UserId

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF( @TranStarted = 1 )
    BEGIN
	SET @TranStarted = 0
	COMMIT TRANSACTION
    END

    RETURN @ErrorCode

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
    	ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode

END

GO
ALTER PROCEDURE dbo.aspnet_Membership_GetUserByUserId
    @UserId               UNIQUEIDENTIFIER,
    @TimeZoneAdjustment   INT,
    @UpdateLastActivity   BIT = 0
AS
BEGIN
    IF ( @UpdateLastActivity = 1 )
    BEGIN
        DECLARE @DateTimeNowUTC DATETIME
        EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT
        UPDATE   dbo.aspnet_Users
        SET      LastActivityDate = @DateTimeNowUTC
        FROM     dbo.aspnet_Users
        WHERE    @UserId = UserId

        IF ( @@ROWCOUNT = 0 ) -- User ID not found
            RETURN -1
    END

    SELECT  m.Email, m.PasswordQuestion, m.Comment, m.IsApproved,
            m.CreateDate, m.LastLoginDate, u.LastActivityDate,
            m.LastPasswordChangedDate, u.UserName, m.IsLockedOut,
            m.LastLockoutDate
    FROM    dbo.aspnet_Users u, dbo.aspnet_Membership m
    WHERE   @UserId = u.UserId AND u.UserId = m.UserId

    IF ( @@ROWCOUNT = 0 ) -- User ID not found
       RETURN -1

    RETURN 0
END

GO
ALTER PROCEDURE dbo.aspnet_Membership_SetPassword
    @ApplicationName  NVARCHAR(256),
    @UserName         NVARCHAR(256),
    @NewPassword      NVARCHAR(128),
    @PasswordSalt     NVARCHAR(128),
    @TimeZoneAdjustment  INT,
    @PasswordFormat   INT = 0
AS
BEGIN
    DECLARE @UserId UNIQUEIDENTIFIER
    SELECT  @UserId = NULL
    SELECT  @UserId = u.UserId
    FROM    dbo.aspnet_Users u, dbo.aspnet_Applications a, dbo.aspnet_Membership m
    WHERE   LoweredUserName = LOWER(@UserName) AND
            u.ApplicationId = a.ApplicationId  AND
            LOWER(@ApplicationName) = a.LoweredApplicationName AND
            u.UserId = m.UserId

    IF (@UserId IS NULL)
        RETURN(1)
    DECLARE @DateTimeNowUTC DATETIME
    EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT

    UPDATE dbo.aspnet_Membership
    SET Password = @NewPassword, PasswordFormat = @PasswordFormat, PasswordSalt = @PasswordSalt,
        LastPasswordChangedDate = @DateTimeNowUTC
    WHERE @UserId = UserId
    RETURN(0)
END

GO

ALTER PROCEDURE dbo.aspnet_Profile_GetProfiles
    @ApplicationName        NVARCHAR(256),
    @ProfileAuthOptions     INT,
    @PageIndex              INT,
    @PageSize               INT,
    @TimeZoneAdjustment     INT,
    @UserNameToMatch        NVARCHAR(256) = NULL,
    @InactiveSinceDate      DATETIME      = NULL
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
        RETURN

    IF ((NOT(@InactiveSinceDate IS NULL)) AND (@InactiveSinceDate > CONVERT(DATETIME, '17540101', 112)) AND  (@InactiveSinceDate < CONVERT(DATETIME, '99980101', 112)))
        SELECT @InactiveSinceDate = DATEADD(n, -@TimeZoneAdjustment, @InactiveSinceDate)

    -- Set the page bounds
    DECLARE @PageLowerBound INT
    DECLARE @PageUpperBound INT
    DECLARE @TotalRecords   INT
    SET @PageLowerBound = @PageSize * @PageIndex
    SET @PageUpperBound = @PageSize - 1 + @PageLowerBound

    -- Create a temp table TO store the select results
    CREATE TABLE #PageIndexForUsers
    (
        IndexId int IDENTITY (0, 1) NOT NULL,
        UserId UNIQUEIDENTIFIER
    )

    -- Insert into our temp table
    INSERT INTO #PageIndexForUsers (UserId)
        SELECT  u.UserId
        FROM    dbo.aspnet_Users u, dbo.aspnet_Profile p
        WHERE   ApplicationId = @ApplicationId
            AND u.UserId = p.UserId
            AND (@InactiveSinceDate IS NULL OR LastActivityDate <= @InactiveSinceDate)
            AND (     (@ProfileAuthOptions = 2)
                   OR (@ProfileAuthOptions = 0 AND IsAnonymous = 1)
                   OR (@ProfileAuthOptions = 1 AND IsAnonymous = 0)
                 )
            AND (@UserNameToMatch IS NULL OR LoweredUserName LIKE LOWER(@UserNameToMatch))
        ORDER BY UserName

    SELECT  u.UserName, u.IsAnonymous, u.LastActivityDate, p.LastUpdatedDate,
            DATALENGTH(p.PropertyNames) + DATALENGTH(p.PropertyValuesString) + DATALENGTH(p.PropertyValuesBinary)
    FROM    dbo.aspnet_Users u, dbo.aspnet_Profile p, #PageIndexForUsers i
    WHERE   u.UserId = p.UserId AND p.UserId = i.UserId AND i.IndexId >= @PageLowerBound AND i.IndexId <= @PageUpperBound

    SELECT COUNT(*)
    FROM   #PageIndexForUsers

    DROP TABLE #PageIndexForUsers
END

GO
ALTER PROCEDURE dbo.aspnet_Membership_UpdateUser
    @ApplicationName      NVARCHAR(256),
    @UserName             NVARCHAR(256),
    @Email                NVARCHAR(256),
    @Comment              NTEXT,
    @IsApproved           BIT,
    @LastLoginDate        DATETIME,
    @LastActivityDate     DATETIME,
    @UniqueEmail          INT,
    @TimeZoneAdjustment   INT
AS
BEGIN
    DECLARE @UserId UNIQUEIDENTIFIER
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @UserId = NULL
    SELECT  @UserId = u.UserId, @ApplicationId = a.ApplicationId
    FROM    dbo.aspnet_Users u, dbo.aspnet_Applications a, dbo.aspnet_Membership m
    WHERE   LoweredUserName = LOWER(@UserName) AND
            u.ApplicationId = a.ApplicationId  AND
            LOWER(@ApplicationName) = a.LoweredApplicationName AND
            u.UserId = m.UserId

    IF (@UserId IS NULL)
        RETURN(1)

    IF (@UniqueEmail = 1)
    BEGIN
        IF (EXISTS (SELECT *
                    FROM  dbo.aspnet_Membership WITH (UPDLOCK, HOLDLOCK)
                    WHERE ApplicationId = @ApplicationId  AND @UserId <> UserId AND LoweredEmail = LOWER(@Email)))
        BEGIN
            RETURN(7)
        END
    END

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
	    BEGIN TRANSACTION
	    SET @TranStarted = 1
    END
    ELSE
	SET @TranStarted = 0

    UPDATE dbo.aspnet_Membership
    SET
         Email            = @Email,
         LoweredEmail     = LOWER(@Email),
         Comment          = @Comment,
         IsApproved       = @IsApproved,
         LastLoginDate    = DATEADD(n, -@TimeZoneAdjustment, @LastLoginDate)
    WHERE
       @UserId = UserId

    IF( @@ERROR <> 0 )
        GOTO Cleanup

    UPDATE dbo.aspnet_Users
    SET
         LastActivityDate = DATEADD(n, -@TimeZoneAdjustment, @LastActivityDate)
    WHERE
       @UserId = UserId

    IF( @@ERROR <> 0 )
        GOTO Cleanup

    IF( @TranStarted = 1 )
    BEGIN
	SET @TranStarted = 0
	COMMIT TRANSACTION
    END

    RETURN 0

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
    	ROLLBACK TRANSACTION
    END

    RETURN -1
END

GO
ALTER PROCEDURE dbo.aspnet_Membership_GetUserByEmail
    @ApplicationName  NVARCHAR(256),
    @Email            NVARCHAR(256)
AS
BEGIN
    IF( @Email IS NULL )
        SELECT  u.UserName
        FROM    dbo.aspnet_Applications a, dbo.aspnet_Users u, dbo.aspnet_Membership m
        WHERE   LOWER(@ApplicationName) = a.LoweredApplicationName AND
                u.ApplicationId = a.ApplicationId    AND
                u.UserId = m.UserId AND
                m.LoweredEmail IS NULL
    ELSE
        SELECT  u.UserName
        FROM    dbo.aspnet_Applications a, dbo.aspnet_Users u, dbo.aspnet_Membership m
        WHERE   LOWER(@ApplicationName) = a.LoweredApplicationName AND
                u.ApplicationId = a.ApplicationId    AND
                u.UserId = m.UserId AND
                LOWER(@Email) = m.LoweredEmail

    IF (@@rowcount = 0)
        RETURN(1)
    RETURN(0)
END

GO

  ALTER VIEW [dbo].[vw_aspnet_Roles]
  AS SELECT [dbo].[aspnet_Roles].[ApplicationId], [dbo].[aspnet_Roles].[RoleId], [dbo].[aspnet_Roles].[RoleName], [dbo].[aspnet_Roles].[LoweredRoleName], [dbo].[aspnet_Roles].[Description]
  FROM [dbo].[aspnet_Roles]
  
GO
ALTER PROCEDURE dbo.aspnet_Membership_GetAllUsers
    @ApplicationName       NVARCHAR(256),
    @PageIndex             INT,
    @PageSize              INT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM dbo.aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
        RETURN 0


    -- Set the page bounds
    DECLARE @PageLowerBound INT
    DECLARE @PageUpperBound INT
    DECLARE @TotalRecords   INT
    SET @PageLowerBound = @PageSize * @PageIndex
    SET @PageUpperBound = @PageSize - 1 + @PageLowerBound

    -- Create a temp table TO store the select results
    CREATE TABLE #PageIndexForUsers
    (
        IndexId int IDENTITY (0, 1) NOT NULL,
        UserId UNIQUEIDENTIFIER
    )

    -- Insert into our temp table
    INSERT INTO #PageIndexForUsers (UserId)
    SELECT u.UserId
    FROM   dbo.aspnet_Membership m, dbo.aspnet_Users u
    WHERE  u.ApplicationId = @ApplicationId AND u.UserId = m.UserId
    ORDER BY u.UserName

    SELECT @TotalRecords = @@ROWCOUNT

    SELECT u.UserName, m.Email, m.PasswordQuestion, m.Comment, m.IsApproved,
            m.CreateDate,
            m.LastLoginDate,
            u.LastActivityDate,
            m.LastPasswordChangedDate,
            u.UserId, m.IsLockedOut,
            m.LastLockoutDate
    FROM   dbo.aspnet_Membership m, dbo.aspnet_Users u, #PageIndexForUsers p
    WHERE  u.UserId = p.UserId AND u.UserId = m.UserId AND
           p.IndexId >= @PageLowerBound AND p.IndexId <= @PageUpperBound
    ORDER BY u.UserName
    RETURN @TotalRecords
END

GO

ALTER PROCEDURE dbo.aspnet_UsersInRoles_RemoveUsersFromRoles
    @ApplicationName  NVARCHAR(256),
    @UserNames        NVARCHAR(4000),
    @RoleNames        NVARCHAR(4000)
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
        RETURN(2)


    DECLARE @TranStarted   BIT
    DECLARE @ErrorCode INT
    SET @ErrorCode   = 0
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
        BEGIN TRANSACTION
        SET @TranStarted = 1
    END
    ELSE
        SET @TranStarted = 0


    DECLARE @RoleId UNIQUEIDENTIFIER
    DECLARE @UserId UNIQUEIDENTIFIER
    DECLARE @UserName     NVARCHAR(256)
    DECLARE @RoleName     NVARCHAR(256)

    DECLARE @CurrentPosU  INT
    DECLARE @NextPosU     INT
    DECLARE @CurrentPosR  INT
    DECLARE @NextPosR     INT

    SELECT  @CurrentPosU = 1

    WHILE(@CurrentPosU <= LEN(@UserNames))
    BEGIN
        SELECT @NextPosU = CHARINDEX(N',', @UserNames,  @CurrentPosU)
        IF (@NextPosU = 0  OR @NextPosU IS NULL)
            SELECT @NextPosU = LEN(@UserNames)+1
        SELECT @UserName = SUBSTRING(@UserNames, @CurrentPosU, @NextPosU - @CurrentPosU)
        SELECT @CurrentPosU = @NextPosU+1

        SELECT @CurrentPosR = 1
        WHILE(@CurrentPosR <= LEN(@RoleNames))
        BEGIN
            SELECT @NextPosR = CHARINDEX(N',', @RoleNames,  @CurrentPosR)
            IF (@NextPosR = 0 OR @NextPosR IS NULL)
                SELECT @NextPosR = LEN(@RoleNames)+1
            SELECT @RoleName = SUBSTRING(@RoleNames, @CurrentPosR, @NextPosR - @CurrentPosR)
            SELECT @CurrentPosR = @NextPosR+1

            SELECT @RoleId = NULL
            SELECT @RoleId = RoleId FROM dbo.aspnet_Roles WHERE LoweredRoleName = LOWER(@RoleName) AND ApplicationId = @ApplicationId
            IF (@RoleId IS NULL)
            BEGIN
                SELECT N'', @RoleName
                SET @ErrorCode = 2
                GOTO Cleanup
            END

            SELECT @UserId = NULL
            SELECT @UserId = UserId FROM dbo.aspnet_Users WHERE LoweredUserName = LOWER(@UserName) AND ApplicationId = @ApplicationId
            IF (@UserId IS NULL)
            BEGIN
                SELECT @UserName, N''
                SET @ErrorCode = 1
                GOTO Cleanup
            END

            IF (NOT(EXISTS(SELECT * FROM dbo.aspnet_UsersInRoles WHERE @UserId = UserId AND @RoleId = RoleId)))
            BEGIN
                SELECT @UserName, @RoleName
                SET @ErrorCode = 3
                GOTO Cleanup
            END
            DELETE FROM dbo.aspnet_UsersInRoles WHERE (UserId = @UserId AND RoleId = @RoleId)
        END
    END

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
        COMMIT TRANSACTION
    END

    RETURN(0)

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
        ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode
END

GO

ALTER PROCEDURE dbo.aspnet_Roles_DeleteRole
    @ApplicationName            NVARCHAR(256),
    @RoleName                   NVARCHAR(256),
    @DeleteOnlyIfRoleIsEmpty    BIT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
        RETURN(1)

    DECLARE @ErrorCode     INT
    SET @ErrorCode = 0

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
        BEGIN TRANSACTION
        SET @TranStarted = 1
    END
    ELSE
        SET @TranStarted = 0

    DECLARE @RoleId   UNIQUEIDENTIFIER
    SELECT  @RoleId = NULL
    SELECT  @RoleId = RoleId FROM dbo.aspnet_Roles WHERE LoweredRoleName = LOWER(@RoleName) AND ApplicationId = @ApplicationId

    IF (@RoleId IS NULL)
    BEGIN
        SELECT @ErrorCode = 1
        GOTO Cleanup
    END
    IF (@DeleteOnlyIfRoleIsEmpty <> 0)
    BEGIN
        IF (EXISTS (SELECT RoleId FROM dbo.aspnet_UsersInRoles  WHERE @RoleId = RoleId))
        BEGIN
            SELECT @ErrorCode = 2
            GOTO Cleanup
        END
    END


    DELETE FROM dbo.aspnet_UsersInRoles  WHERE @RoleId = RoleId

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    DELETE FROM dbo.aspnet_Roles WHERE @RoleId = RoleId  AND ApplicationId = @ApplicationId

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
        COMMIT TRANSACTION
    END

    RETURN(0)

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
        ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode
END

GO

ALTER PROCEDURE dbo.aspnet_Profile_GetNumberOfInactiveProfiles
    @ApplicationName        NVARCHAR(256),
    @ProfileAuthOptions     INT,
    @InactiveSinceDate      DATETIME,
    @TimeZoneAdjustment     INT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
    BEGIN
        SELECT 0
        RETURN
    END

    IF (@InactiveSinceDate > CONVERT(DATETIME, '17540101', 112) AND  @InactiveSinceDate < CONVERT(DATETIME,'99980101', 112))
        SELECT @InactiveSinceDate = DATEADD(n, -@TimeZoneAdjustment, @InactiveSinceDate)

    SELECT  COUNT(*)
    FROM    dbo.aspnet_Users u, dbo.aspnet_Profile p
    WHERE   ApplicationId = @ApplicationId
        AND u.UserId = p.UserId
        AND (LastActivityDate <= @InactiveSinceDate)
        AND (
                (@ProfileAuthOptions = 2)
                OR (@ProfileAuthOptions = 0 AND IsAnonymous = 1)
                OR (@ProfileAuthOptions = 1 AND IsAnonymous = 0)
            )
END

GO
ALTER PROCEDURE dbo.aspnet_Membership_FindUsersByEmail
    @ApplicationName       NVARCHAR(256),
    @EmailToMatch          NVARCHAR(256),
    @PageIndex             INT,
    @PageSize              INT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM dbo.aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
        RETURN 0

    -- Set the page bounds
    DECLARE @PageLowerBound INT
    DECLARE @PageUpperBound INT
    DECLARE @TotalRecords   INT
    SET @PageLowerBound = @PageSize * @PageIndex
    SET @PageUpperBound = @PageSize - 1 + @PageLowerBound

    -- Create a temp table TO store the select results
    CREATE TABLE #PageIndexForUsers
    (
        IndexId int IDENTITY (0, 1) NOT NULL,
        UserId UNIQUEIDENTIFIER
    )

    -- Insert into our temp table
    IF( @EmailToMatch IS NULL )
        INSERT INTO #PageIndexForUsers (UserId)
            SELECT u.UserId
            FROM   dbo.aspnet_Users u, dbo.aspnet_Membership m
            WHERE  u.ApplicationId = @ApplicationId AND m.UserId = u.UserId AND m.Email IS NULL
            ORDER BY m.LoweredEmail
    ELSE
        INSERT INTO #PageIndexForUsers (UserId)
            SELECT u.UserId
            FROM   dbo.aspnet_Users u, dbo.aspnet_Membership m
            WHERE  u.ApplicationId = @ApplicationId AND m.UserId = u.UserId AND m.LoweredEmail LIKE LOWER(@EmailToMatch)
            ORDER BY m.LoweredEmail

    SELECT  u.UserName, m.Email, m.PasswordQuestion, m.Comment, m.IsApproved,
            m.CreateDate,
            m.LastLoginDate,
            u.LastActivityDate,
            m.LastPasswordChangedDate,
            u.UserId, m.IsLockedOut,
            m.LastLockoutDate
    FROM   dbo.aspnet_Membership m, dbo.aspnet_Users u, #PageIndexForUsers p
    WHERE  u.UserId = p.UserId AND u.UserId = m.UserId AND
           p.IndexId >= @PageLowerBound AND p.IndexId <= @PageUpperBound
    ORDER BY m.LoweredEmail

    SELECT  @TotalRecords = COUNT(*)
    FROM    #PageIndexForUsers
    RETURN @TotalRecords
END

GO
ALTER PROCEDURE dbo.aspnet_Membership_ResetPassword
    @ApplicationName             NVARCHAR(256),
    @UserName                    NVARCHAR(256),
    @NewPassword                 NVARCHAR(128),
    @MaxInvalidPasswordAttempts  INT,
    @PasswordAttemptWindow       INT,
    @PasswordSalt                NVARCHAR(128),
    @TimeZoneAdjustment          INT,
    @PasswordFormat              INT = 0,
    @PasswordAnswer              NVARCHAR(128) = NULL
AS
BEGIN
    DECLARE @IsLockedOut                            BIT
    DECLARE @LastLockoutDate                        DATETIME
    DECLARE @FailedPasswordAttemptCount             INT
    DECLARE @FailedPasswordAttemptWindowStart       DATETIME
    DECLARE @FailedPasswordAnswerAttemptCount       INT
    DECLARE @FailedPasswordAnswerAttemptWindowStart DATETIME

    DECLARE @UserId                                 UNIQUEIDENTIFIER
    SET     @UserId = NULL

    DECLARE @ErrorCode     INT
    SET @ErrorCode = 0

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
	    BEGIN TRANSACTION
	    SET @TranStarted = 1
    END
    ELSE
    	SET @TranStarted = 0

    SELECT  @UserId = u.UserId
    FROM    dbo.aspnet_Users u, dbo.aspnet_Applications a, dbo.aspnet_Membership m
    WHERE   LoweredUserName = LOWER(@UserName) AND
            u.ApplicationId = a.ApplicationId  AND
            LoweredApplicationName = a.LoweredApplicationName AND
            u.UserId = m.UserId

    IF ( @UserId IS NULL )
    BEGIN
        SET @ErrorCode = 1
        GOTO Cleanup
    END

    SELECT @IsLockedOut = IsLockedOut,
           @LastLockoutDate = LastLockoutDate,
           @FailedPasswordAttemptCount = FailedPasswordAttemptCount,
           @FailedPasswordAttemptWindowStart = FailedPasswordAttemptWindowStart,
           @FailedPasswordAnswerAttemptCount = FailedPasswordAnswerAttemptCount,
           @FailedPasswordAnswerAttemptWindowStart = FailedPasswordAnswerAttemptWindowStart
    FROM dbo.aspnet_Membership WITH ( UPDLOCK )
    WHERE @UserId = UserId

    IF( @IsLockedOut = 1 )
    BEGIN
        SET @ErrorCode = 99
        GOTO Cleanup
    END
    DECLARE @DateTimeNowUTC DATETIME
    EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT

    UPDATE dbo.aspnet_Membership
    SET    Password = @NewPassword,
           LastPasswordChangedDate = @DateTimeNowUTC,
           PasswordFormat = @PasswordFormat,
           PasswordSalt = @PasswordSalt
    WHERE  @UserId = UserId AND
           ( ( @PasswordAnswer IS NULL ) OR ( LOWER( PasswordAnswer ) = LOWER( @PasswordAnswer ) ) )

    IF ( @@ROWCOUNT = 0 )
        BEGIN
            IF( @DateTimeNowUTC > DATEADD( minute, @PasswordAttemptWindow, @FailedPasswordAnswerAttemptWindowStart ) )
            BEGIN
                SET @FailedPasswordAnswerAttemptWindowStart = @DateTimeNowUTC
                SET @FailedPasswordAnswerAttemptCount = 1
            END
            ELSE
            BEGIN
                SET @FailedPasswordAnswerAttemptWindowStart = @DateTimeNowUTC
                SET @FailedPasswordAnswerAttemptCount = @FailedPasswordAnswerAttemptCount + 1
            END
            
            BEGIN
                IF( @FailedPasswordAnswerAttemptCount >= @MaxInvalidPasswordAttempts )
                BEGIN
                    SET @IsLockedOut = 1
                    SET @LastLockoutDate = @DateTimeNowUTC
                END
            END

            SET @ErrorCode = 3
        END
    ELSE
        BEGIN
            IF( @FailedPasswordAnswerAttemptCount > 0 )
            BEGIN
                SET @FailedPasswordAnswerAttemptCount = 0
                SET @FailedPasswordAnswerAttemptWindowStart = CONVERT( DATETIME, '17540101', 112 )
            END
        END

    IF( NOT ( @PasswordAnswer IS NULL ) )
    BEGIN
        UPDATE dbo.aspnet_Membership
        SET IsLockedOut = @IsLockedOut, LastLockoutDate = @LastLockoutDate,
            FailedPasswordAttemptCount = @FailedPasswordAttemptCount,
            FailedPasswordAttemptWindowStart = @FailedPasswordAttemptWindowStart,
            FailedPasswordAnswerAttemptCount = @FailedPasswordAnswerAttemptCount,
            FailedPasswordAnswerAttemptWindowStart = @FailedPasswordAnswerAttemptWindowStart
        WHERE @UserId = UserId

        IF( @@ERROR <> 0 )
        BEGIN
            SET @ErrorCode = -1
            GOTO Cleanup
        END
    END

    IF( @TranStarted = 1 )
    BEGIN
	SET @TranStarted = 0
	COMMIT TRANSACTION
    END

    RETURN @ErrorCode

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
    	ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode

END

GO
ALTER PROCEDURE dbo.aspnet_Membership_GetPasswordWithFormat
    @ApplicationName                NVARCHAR(256),
    @UserName                       NVARCHAR(256)
AS
BEGIN
    DECLARE @Password                               NVARCHAR(128)
    DECLARE @PasswordFormat                         INT
    DECLARE @PasswordSalt                           NVARCHAR(128)
    DECLARE @IsLockedOut                            BIT
    DECLARE @FailedPasswordAttemptCount             INT
    DECLARE @FailedPasswordAnswerAttemptCount       INT
    DECLARE @IsApproved                             BIT

    SELECT  @Password = m.Password,
            @PasswordFormat = m.PasswordFormat,
            @PasswordSalt = m.PasswordSalt,
            @IsLockedOut = m.IsLockedOut,
            @FailedPasswordAttemptCount = m.FailedPasswordAttemptCount,
            @FailedPasswordAnswerAttemptCount = m.FailedPasswordAnswerAttemptCount,
            @IsApproved = m.IsApproved
    FROM    dbo.aspnet_Applications a, dbo.aspnet_Users u, dbo.aspnet_Membership m
    WHERE   LOWER(@ApplicationName) = a.LoweredApplicationName AND
            u.ApplicationId = a.ApplicationId    AND
            u.UserId = m.UserId AND
            LOWER(@UserName) = u.LoweredUserName

    IF( @@rowcount = 0 )
        RETURN 1

    IF( @IsLockedOut = 1 )
        RETURN 99

    SELECT @Password,
           @PasswordFormat,
           @PasswordSalt,
           @FailedPasswordAttemptCount,
           @FailedPasswordAnswerAttemptCount,
           @IsApproved

    RETURN 0
END

GO

ALTER PROCEDURE dbo.aspnet_Profile_GetProperties
    @ApplicationName      NVARCHAR(256),
    @UserName             NVARCHAR(256),
    @TimeZoneAdjustment   INT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM dbo.aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
        RETURN

    DECLARE @UserId UNIQUEIDENTIFIER
    SELECT  @UserId = NULL
    DECLARE @DateTimeNowUTC DATETIME
    EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT

    SELECT @UserId = UserId
    FROM   dbo.aspnet_Users
    WHERE  ApplicationId = @ApplicationId AND LoweredUserName = LOWER(@UserName)

    IF (@UserId IS NULL)
        RETURN
    SELECT TOP 1 PropertyNames, PropertyValuesString, PropertyValuesBinary
    FROM         dbo.aspnet_Profile
    WHERE        UserId = @UserId

    IF (@@ROWCOUNT > 0)
    BEGIN
        UPDATE dbo.aspnet_Users
        SET    LastActivityDate=@DateTimeNowUTC
        WHERE  UserId = @UserId
    END
END

GO
ALTER PROCEDURE dbo.aspnet_Membership_UpdateLastLoginAndActivityDates
    @ApplicationName          NVARCHAR(256),
    @UserName                 NVARCHAR(256),
    @TimeZoneAdjustment       INT
AS
BEGIN
    DECLARE @UserId UNIQUEIDENTIFIER
    SELECT  @UserId = NULL
    SELECT  @UserId = u.UserId
    FROM    dbo.aspnet_Membership m, dbo.aspnet_Users u, dbo.aspnet_Applications a
    WHERE   LoweredUserName = LOWER(@UserName) AND
            u.ApplicationId = a.ApplicationId  AND
            LOWER(@ApplicationName) = a.LoweredApplicationName AND
            u.UserId = m.UserId
    IF (@UserId IS NULL)
    BEGIN
        RETURN
    END

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
	    BEGIN TRANSACTION
	    SET @TranStarted = 1
    END
    ELSE
	  SET @TranStarted = 0

    DECLARE @DateTimeNowUTC DATETIME
    EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT

    UPDATE  dbo.aspnet_Membership
    SET     LastLoginDate = @DateTimeNowUTC
    WHERE   UserId = @UserId

    IF( @@ERROR <> 0 )
        GOTO Cleanup

    UPDATE  dbo.aspnet_Users
    SET     LastActivityDate = @DateTimeNowUTC
    WHERE   @UserId = UserId

    IF( @@ERROR <> 0 )
        GOTO Cleanup

    IF( @TranStarted = 1 )
    BEGIN
	SET @TranStarted = 0
	COMMIT TRANSACTION
    END

    RETURN

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
	    ROLLBACK TRANSACTION
    END

    RETURN -1

END

GO
ALTER PROCEDURE dbo.aspnet_Membership_CreateUser
    @ApplicationName                        NVARCHAR(256),
    @UserName                               NVARCHAR(256),
    @Password                               NVARCHAR(128),
    @PasswordSalt                           NVARCHAR(128),
    @Email                                  NVARCHAR(256),
    @PasswordQuestion                       NVARCHAR(256),
    @PasswordAnswer                         NVARCHAR(128),
    @IsApproved                             BIT,
    @TimeZoneAdjustment                     INT,
    @CreateDate                             DATETIME = NULL,
    @UniqueEmail                            INT      = 0,
    @PasswordFormat                         INT      = 0,
    @UserId                                 UNIQUEIDENTIFIER OUTPUT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL

    DECLARE @NewUserId UNIQUEIDENTIFIER
    SELECT @NewUserId = NULL

    DECLARE @IsLockedOut BIT
    SET @IsLockedOut = 0

    DECLARE @LastLockoutDate  DATETIME
    SET @LastLockoutDate = CONVERT( DATETIME, '17540101', 112 )

    DECLARE @FailedPasswordAttemptCount INT
    SET @FailedPasswordAttemptCount = 0

    DECLARE @FailedPasswordAttemptWindowStart  DATETIME
    SET @FailedPasswordAttemptWindowStart = CONVERT( DATETIME, '17540101', 112 )

    DECLARE @FailedPasswordAnswerAttemptCount INT
    SET @FailedPasswordAnswerAttemptCount = 0

    DECLARE @FailedPasswordAnswerAttemptWindowStart  DATETIME
    SET @FailedPasswordAnswerAttemptWindowStart = CONVERT( DATETIME, '17540101', 112 )

    DECLARE @NewUserCreated BIT
    DECLARE @ReturnValue   INT
    SET @ReturnValue = 0

    DECLARE @ErrorCode     INT
    SET @ErrorCode = 0

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
	    BEGIN TRANSACTION
	    SET @TranStarted = 1
    END
    ELSE
    	SET @TranStarted = 0

    EXEC dbo.aspnet_Applications_CreateApplication @ApplicationName, @ApplicationId OUTPUT

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF (@CreateDate IS NULL)
        EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @CreateDate OUTPUT
    ELSE
        SELECT  @CreateDate = DATEADD(n, -@TimeZoneAdjustment, @CreateDate) -- switch TO UTC time

    SELECT  @NewUserId = UserId FROM dbo.aspnet_Users WHERE LOWER(@UserName) = LoweredUserName AND @ApplicationId = ApplicationId
    IF ( @NewUserId IS NULL )
    BEGIN
        SET @NewUserId = @UserId
        EXEC @ReturnValue = dbo.aspnet_Users_CreateUser @ApplicationId, @UserName, 0, @CreateDate, @NewUserId OUTPUT
        SET @NewUserCreated = 1
    END
    ELSE
    BEGIN
        SET @NewUserCreated = 0
        IF( @NewUserId <> @UserId AND @UserId IS NOT NULL )
        BEGIN
            SET @ErrorCode = 6
            GOTO Cleanup
        END
    END

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF( @ReturnValue = -1 )
    BEGIN
        SET @ErrorCode = 10
        GOTO Cleanup
    END

    IF ( EXISTS ( SELECT UserId
                  FROM   dbo.aspnet_Membership
                  WHERE  @NewUserId = UserId ) )
    BEGIN
        SET @ErrorCode = 6
        GOTO Cleanup
    END

    SET @UserId = @NewUserId

    IF (@UniqueEmail = 1)
    BEGIN
        IF (EXISTS (SELECT *
                    FROM  dbo.aspnet_Membership m WITH ( UPDLOCK, HOLDLOCK )
                    WHERE ApplicationId = @ApplicationId AND LoweredEmail = LOWER(@Email)))
        BEGIN
            SET @ErrorCode = 7
            GOTO Cleanup
        END
    END

    INSERT INTO dbo.aspnet_Membership
                ( ApplicationId,
                  UserId,
                  Password,
                  PasswordSalt,
                  Email,
                  LoweredEmail,
                  PasswordQuestion,
                  PasswordAnswer,
                  PasswordFormat,
                  IsApproved,
                  IsLockedOut,
                  CreateDate,
                  LastLoginDate,
                  LastPasswordChangedDate,
                  LastLockoutDate,
                  FailedPasswordAttemptCount,
                  FailedPasswordAttemptWindowStart,
                  FailedPasswordAnswerAttemptCount,
                  FailedPasswordAnswerAttemptWindowStart )
         VALUES ( @ApplicationId,
                  @UserId,
                  @Password,
                  @PasswordSalt,
                  @Email,
                  LOWER(@Email),
                  @PasswordQuestion,
                  @PasswordAnswer,
                  @PasswordFormat,
                  @IsApproved,
                  @IsLockedOut,
                  @CreateDate,
                  @CreateDate,
                  @CreateDate,
                  @LastLockoutDate,
                  @FailedPasswordAttemptCount,
                  @FailedPasswordAttemptWindowStart,
                  @FailedPasswordAnswerAttemptCount,
                  @FailedPasswordAnswerAttemptWindowStart )

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF (@NewUserCreated = 0)
    BEGIN
        UPDATE dbo.aspnet_Users
        SET    LastActivityDate = @CreateDate
        WHERE  @UserId = UserId
        IF( @@ERROR <> 0 )
        BEGIN
            SET @ErrorCode = -1
            GOTO Cleanup
        END
    END

    SELECT @CreateDate = DATEADD( n, @TimeZoneAdjustment, @CreateDate )

    IF( @TranStarted = 1 )
    BEGIN
	    SET @TranStarted = 0
	    COMMIT TRANSACTION
    END

    RETURN 0

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
    	ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode

END

GO
ALTER PROCEDURE dbo.aspnet_Membership_GetPassword
    @ApplicationName                NVARCHAR(256),
    @UserName                       NVARCHAR(256),
    @MaxInvalidPasswordAttempts     INT,
    @PasswordAttemptWindow          INT,
    @TimeZoneAdjustment             INT,
    @PasswordAnswer                 NVARCHAR(128) = NULL
AS
BEGIN
    DECLARE @UserId                                 UNIQUEIDENTIFIER
    DECLARE @PasswordFormat                         INT
    DECLARE @Password                               NVARCHAR(128)
    DECLARE @passAns                                NVARCHAR(128)
    DECLARE @IsLockedOut                            BIT
    DECLARE @LastLockoutDate                        DATETIME
    DECLARE @FailedPasswordAttemptCount             INT
    DECLARE @FailedPasswordAttemptWindowStart       DATETIME
    DECLARE @FailedPasswordAnswerAttemptCount       INT
    DECLARE @FailedPasswordAnswerAttemptWindowStart DATETIME

    DECLARE @ErrorCode     INT
    SET @ErrorCode = 0

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
	    BEGIN TRANSACTION
	    SET @TranStarted = 1
    END
    ELSE
    	SET @TranStarted = 0

    DECLARE @DateTimeNowUTC DATETIME
    EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT

    SELECT  @UserId = u.UserId,
            @Password = m.Password,
            @passAns = m.PasswordAnswer,
            @PasswordFormat = m.PasswordFormat,
            @IsLockedOut = m.IsLockedOut,
            @LastLockoutDate = m.LastLockoutDate,
            @FailedPasswordAttemptCount = m.FailedPasswordAttemptCount,
            @FailedPasswordAttemptWindowStart = m.FailedPasswordAttemptWindowStart,
            @FailedPasswordAnswerAttemptCount = m.FailedPasswordAnswerAttemptCount,
            @FailedPasswordAnswerAttemptWindowStart = m.FailedPasswordAnswerAttemptWindowStart
    FROM    dbo.aspnet_Applications a, dbo.aspnet_Users u, dbo.aspnet_Membership m WITH ( UPDLOCK )
    WHERE   LOWER(@ApplicationName) = a.LoweredApplicationName AND
            u.ApplicationId = a.ApplicationId    AND
            u.UserId = m.UserId AND
            LOWER(@UserName) = u.LoweredUserName

    IF ( @@rowcount = 0 )
    BEGIN
        SET @ErrorCode = 1
        GOTO Cleanup
    END

    IF( @IsLockedOut = 1 )
    BEGIN
        SET @ErrorCode = 99
        GOTO Cleanup
    END

    IF ( NOT( @PasswordAnswer IS NULL ) )
    BEGIN
        IF( ( @passAns IS NULL ) OR ( LOWER( @passAns ) <> LOWER( @PasswordAnswer ) ) )
        BEGIN
            IF( @DateTimeNowUTC > DATEADD( minute, @PasswordAttemptWindow, @FailedPasswordAnswerAttemptWindowStart ) )
            BEGIN
                SET @FailedPasswordAnswerAttemptWindowStart = @DateTimeNowUTC
                SET @FailedPasswordAnswerAttemptCount = 1
            END
            ELSE
            BEGIN
                SET @FailedPasswordAnswerAttemptCount = @FailedPasswordAnswerAttemptCount + 1
                SET @FailedPasswordAnswerAttemptWindowStart = @DateTimeNowUTC
            END

            BEGIN
                IF( @FailedPasswordAnswerAttemptCount >= @MaxInvalidPasswordAttempts )
                BEGIN
                    SET @IsLockedOut = 1
                    SET @LastLockoutDate = @DateTimeNowUTC
                END
            END

            SET @ErrorCode = 3
        END
        ELSE
        BEGIN
            IF( @FailedPasswordAnswerAttemptCount > 0 )
            BEGIN
                SET @FailedPasswordAnswerAttemptCount = 0
                SET @FailedPasswordAnswerAttemptWindowStart = CONVERT( DATETIME, '17540101', 112 )
            END
        END

        UPDATE dbo.aspnet_Membership
        SET IsLockedOut = @IsLockedOut, LastLockoutDate = @LastLockoutDate,
            FailedPasswordAttemptCount = @FailedPasswordAttemptCount,
            FailedPasswordAttemptWindowStart = @FailedPasswordAttemptWindowStart,
            FailedPasswordAnswerAttemptCount = @FailedPasswordAnswerAttemptCount,
            FailedPasswordAnswerAttemptWindowStart = @FailedPasswordAnswerAttemptWindowStart
        WHERE @UserId = UserId

        IF( @@ERROR <> 0 )
        BEGIN
            SET @ErrorCode = -1
            GOTO Cleanup
        END
    END

    IF( @TranStarted = 1 )
    BEGIN
	SET @TranStarted = 0
	COMMIT TRANSACTION
    END

    IF( @ErrorCode = 0 )
        SELECT @Password, @PasswordFormat

    RETURN @ErrorCode

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
    	ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode

END

GO
ALTER PROCEDURE dbo.aspnet_Membership_GetUserByName
    @ApplicationName      NVARCHAR(256),
    @UserName             NVARCHAR(256),
    @TimeZoneAdjustment   INT,
    @UpdateLastActivity   BIT = 0
AS
BEGIN
    IF (@UpdateLastActivity = 1)
    BEGIN
        DECLARE @DateTimeNowUTC DATETIME
        EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT
        UPDATE   dbo.aspnet_Users
        SET      LastActivityDate = @DateTimeNowUTC
        FROM     dbo.aspnet_Applications a, dbo.aspnet_Users u
        WHERE    LOWER(@ApplicationName) = a.LoweredApplicationName AND
                 u.ApplicationId = a.ApplicationId    AND
                 u.LoweredUserName = LOWER(@UserName)

        IF (@@ROWCOUNT = 0) -- Username not found
            RETURN -1
    END

    SELECT  m.Email, m.PasswordQuestion, m.Comment, m.IsApproved,
            m.CreateDate, m.LastLoginDate, u.LastActivityDate, m.LastPasswordChangedDate,
            u.UserId, m.IsLockedOut,m.LastLockoutDate
    FROM    dbo.aspnet_Applications a, dbo.aspnet_Users u, dbo.aspnet_Membership m
    WHERE   LOWER(@ApplicationName) = a.LoweredApplicationName AND
            u.ApplicationId = a.ApplicationId    AND
            u.LoweredUserName = LOWER(@UserName) AND
            u.UserId = m.UserId
    IF (@@ROWCOUNT = 0) -- Username not found
       RETURN -1

    RETURN 0
END

GO
ALTER PROCEDURE dbo.aspnet_Membership_FindUsersByName
    @ApplicationName       NVARCHAR(256),
    @UserNameToMatch       NVARCHAR(256),
    @PageIndex             INT,
    @PageSize              INT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM dbo.aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
        RETURN 0

    -- Set the page bounds
    DECLARE @PageLowerBound INT
    DECLARE @PageUpperBound INT
    DECLARE @TotalRecords   INT
    SET @PageLowerBound = @PageSize * @PageIndex
    SET @PageUpperBound = @PageSize - 1 + @PageLowerBound

    -- Create a temp table TO store the select results
    CREATE TABLE #PageIndexForUsers
    (
        IndexId int IDENTITY (0, 1) NOT NULL,
        UserId UNIQUEIDENTIFIER
    )

    -- Insert into our temp table
    INSERT INTO #PageIndexForUsers (UserId)
        SELECT u.UserId
        FROM   dbo.aspnet_Users u, dbo.aspnet_Membership m
        WHERE  u.ApplicationId = @ApplicationId AND m.UserId = u.UserId AND u.LoweredUserName LIKE LOWER(@UserNameToMatch)
        ORDER BY u.UserName


    SELECT  u.UserName, m.Email, m.PasswordQuestion, m.Comment, m.IsApproved,
            m.CreateDate,
            m.LastLoginDate,
            u.LastActivityDate,
            m.LastPasswordChangedDate,
            u.UserId, m.IsLockedOut,
            m.LastLockoutDate
    FROM   dbo.aspnet_Membership m, dbo.aspnet_Users u, #PageIndexForUsers p
    WHERE  u.UserId = p.UserId AND u.UserId = m.UserId AND
           p.IndexId >= @PageLowerBound AND p.IndexId <= @PageUpperBound
    ORDER BY u.UserName

    SELECT  @TotalRecords = COUNT(*)
    FROM    #PageIndexForUsers
    RETURN @TotalRecords
END

GO
ALTER PROCEDURE dbo.aspnet_Membership_UnlockUser
    @ApplicationName                         NVARCHAR(256),
    @UserName                                NVARCHAR(256)
AS
BEGIN
    DECLARE @UserId UNIQUEIDENTIFIER
    SELECT  @UserId = NULL
    SELECT  @UserId = u.UserId
    FROM    dbo.aspnet_Users u, dbo.aspnet_Applications a, dbo.aspnet_Membership m
    WHERE   LoweredUserName = LOWER(@UserName) AND
            u.ApplicationId = a.ApplicationId  AND
            LOWER(@ApplicationName) = a.LoweredApplicationName AND
            u.UserId = m.UserId

    IF ( @UserId IS NULL )
        RETURN 1

    UPDATE dbo.aspnet_Membership
    SET IsLockedOut = 0,
        FailedPasswordAttemptCount = 0,
        FailedPasswordAttemptWindowStart = CONVERT( DATETIME, '17540101', 112 ),
        FailedPasswordAnswerAttemptCount = 0,
        FailedPasswordAnswerAttemptWindowStart = CONVERT( DATETIME, '17540101', 112 ),
        LastLockoutDate = CONVERT( DATETIME, '17540101', 112 )
    WHERE @UserId = UserId

    RETURN 0
END

GO

ALTER PROCEDURE dbo.aspnet_Profile_DeleteProfiles
    @ApplicationName        NVARCHAR(256),
    @UserNames              NVARCHAR(4000)
AS
BEGIN
    DECLARE @UserName     NVARCHAR(256)
    DECLARE @CurrentPos   INT
    DECLARE @NextPos      INT
    DECLARE @NumDeleted   INT
    DECLARE @DeletedUser  INT
    DECLARE @TranStarted  BIT
    DECLARE @ErrorCode    INT

    SET @ErrorCode = 0
    SET @CurrentPos = 1
    SET @NumDeleted = 0
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
        BEGIN TRANSACTION
        SET @TranStarted = 1
    END
    ELSE
    	SET @TranStarted = 0

    WHILE (@CurrentPos <= LEN(@UserNames))
    BEGIN
        SELECT @NextPos = CHARINDEX(N',', @UserNames,  @CurrentPos)
        IF (@NextPos = 0 OR @NextPos IS NULL)
            SELECT @NextPos = LEN(@UserNames) + 1

        SELECT @UserName = SUBSTRING(@UserNames, @CurrentPos, @NextPos - @CurrentPos)
        SELECT @CurrentPos = @NextPos+1

        IF (LEN(@UserName) > 0)
        BEGIN
            SELECT @DeletedUser = 0
            EXEC dbo.aspnet_Users_DeleteUser @ApplicationName, @UserName, 4, @DeletedUser OUTPUT
            IF( @@ERROR <> 0 )
            BEGIN
                SET @ErrorCode = -1
                GOTO Cleanup
            END
            IF (@DeletedUser <> 0)
                SELECT @NumDeleted = @NumDeleted + 1
        END
    END
    SELECT @NumDeleted
    IF (@TranStarted = 1)
    BEGIN
    	SET @TranStarted = 0
    	COMMIT TRANSACTION
    END
    SET @TranStarted = 0

    RETURN 0

Cleanup:
    IF (@TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
    	ROLLBACK TRANSACTION
    END
    RETURN @ErrorCode
END

GO

  ALTER VIEW [dbo].[vw_aspnet_Users]
  AS SELECT [dbo].[aspnet_Users].[ApplicationId], [dbo].[aspnet_Users].[UserId], [dbo].[aspnet_Users].[UserName], [dbo].[aspnet_Users].[LoweredUserName], [dbo].[aspnet_Users].[MobileAlias], [dbo].[aspnet_Users].[IsAnonymous], [dbo].[aspnet_Users].[LastActivityDate]
  FROM [dbo].[aspnet_Users]
  
GO
ALTER PROCEDURE dbo.aspnet_Roles_CreateRole
    @ApplicationName  NVARCHAR(256),
    @RoleName         NVARCHAR(256)
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL

    DECLARE @ErrorCode     INT
    SET @ErrorCode = 0

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
        BEGIN TRANSACTION
        SET @TranStarted = 1
    END
    ELSE
        SET @TranStarted = 0

    EXEC dbo.aspnet_Applications_CreateApplication @ApplicationName, @ApplicationId OUTPUT

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF (EXISTS(SELECT RoleId FROM dbo.aspnet_Roles WHERE LoweredRoleName = LOWER(@RoleName) AND ApplicationId = @ApplicationId))
    BEGIN
        SET @ErrorCode = 1
        GOTO Cleanup
    END

    INSERT INTO dbo.aspnet_Roles
                (ApplicationId, RoleName, LoweredRoleName)
         VALUES (@ApplicationId, @RoleName, LOWER(@RoleName))

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
        COMMIT TRANSACTION
    END

    RETURN(0)

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
        ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode

END

GO

ALTER PROCEDURE dbo.aspnet_Profile_SetProperties
    @ApplicationName        NVARCHAR(256),
    @PropertyNames          NTEXT,
    @PropertyValuesString   NTEXT,
    @PropertyValuesBinary   IMAGE,
    @UserName               NVARCHAR(256),
    @IsUserAnonymous        BIT,
    @TimeZoneAdjustment     INT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL

    DECLARE @ErrorCode     INT
    SET @ErrorCode = 0

    DECLARE @TranStarted   BIT
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
       BEGIN TRANSACTION
       SET @TranStarted = 1
    END
    ELSE
    	SET @TranStarted = 0

    EXEC dbo.aspnet_Applications_CreateApplication @ApplicationName, @ApplicationId OUTPUT

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    DECLARE @DateTimeNowUTC DATETIME
    EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT
    DECLARE @UserId UNIQUEIDENTIFIER
    DECLARE @LastActivityDate DATETIME
    SELECT  @UserId = NULL
    SELECT @LastActivityDate = @DateTimeNowUTC

    SELECT @UserId = UserId
    FROM   dbo.aspnet_Users
    WHERE  ApplicationId = @ApplicationId AND LoweredUserName = LOWER(@UserName)
    IF (@UserId IS NULL)
        EXEC dbo.aspnet_Users_CreateUser @ApplicationId, @UserName, @IsUserAnonymous, @LastActivityDate, @UserId OUTPUT

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF (EXISTS( SELECT *
               FROM   dbo.aspnet_Profile
               WHERE  UserId = @UserId))
        UPDATE dbo.aspnet_Profile
        SET    PropertyNames=@PropertyNames, PropertyValuesString = @PropertyValuesString,
               PropertyValuesBinary = @PropertyValuesBinary, LastUpdatedDate=@DateTimeNowUTC
        WHERE  UserId = @UserId
    ELSE
        INSERT INTO dbo.aspnet_Profile(UserId, PropertyNames, PropertyValuesString, PropertyValuesBinary, LastUpdatedDate)
             VALUES (@UserId, @PropertyNames, @PropertyValuesString, @PropertyValuesBinary, @DateTimeNowUTC)

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    UPDATE dbo.aspnet_Users
    SET    LastActivityDate=@DateTimeNowUTC
    WHERE  UserId = @UserId

    IF( @@ERROR <> 0 )
    BEGIN
        SET @ErrorCode = -1
        GOTO Cleanup
    END

    IF( @TranStarted = 1 )
    BEGIN
    	SET @TranStarted = 0
    	COMMIT TRANSACTION
    END

    RETURN 0

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
    	ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode

END

GO

ALTER PROCEDURE dbo.aspnet_UsersInRoles_AddUsersToRoles
    @ApplicationName        NVARCHAR(256),
    @UserNames              NVARCHAR(4000),
    @RoleNames              NVARCHAR(4000),
    @TimeZoneAdjustment     INT
AS
BEGIN
    DECLARE @ApplicationId UNIQUEIDENTIFIER
    SELECT  @ApplicationId = NULL
    SELECT  @ApplicationId = ApplicationId FROM aspnet_Applications WHERE LOWER(@ApplicationName) = LoweredApplicationName
    IF (@ApplicationId IS NULL)
        RETURN(2)


    DECLARE @TranStarted   BIT
    DECLARE @ErrorCode INT
    SET @ErrorCode   = 0
    SET @TranStarted = 0

    IF( @@TRANCOUNT = 0 )
    BEGIN
        BEGIN TRANSACTION
        SET @TranStarted = 1
    END
    ELSE
        SET @TranStarted = 0

    DECLARE @RoleId UNIQUEIDENTIFIER
    DECLARE @UserId UNIQUEIDENTIFIER
    DECLARE @UserName     NVARCHAR(256)
    DECLARE @RoleName     NVARCHAR(256)

    DECLARE @CurrentPosU  INT
    DECLARE @NextPosU     INT
    DECLARE @CurrentPosR  INT
    DECLARE @NextPosR     INT

    SELECT  @CurrentPosU = 1

    WHILE(@CurrentPosU <= LEN(@UserNames))
    BEGIN
        SELECT @NextPosU = CHARINDEX(N',', @UserNames,  @CurrentPosU)
        IF (@NextPosU = 0 OR @NextPosU IS NULL)
            SELECT @NextPosU = LEN(@UserNames) + 1

        SELECT @UserName = SUBSTRING(@UserNames, @CurrentPosU, @NextPosU - @CurrentPosU)
        SELECT @CurrentPosU = @NextPosU+1

        SELECT @CurrentPosR = 1
        WHILE(@CurrentPosR <= LEN(@RoleNames))
        BEGIN
            SELECT @NextPosR = CHARINDEX(N',', @RoleNames,  @CurrentPosR)
            IF (@NextPosR = 0 OR @NextPosR IS NULL)
                SELECT @NextPosR = LEN(@RoleNames) + 1
            SELECT @RoleName = SUBSTRING(@RoleNames, @CurrentPosR, @NextPosR - @CurrentPosR)
            SELECT @CurrentPosR = @NextPosR+1


            SELECT @RoleId = NULL
            SELECT @RoleId = RoleId FROM dbo.aspnet_Roles WHERE LoweredRoleName = LOWER(@RoleName) AND ApplicationId = @ApplicationId
            IF (@RoleId IS NULL)
            BEGIN
                SELECT @RoleName
                SET @ErrorCode = 2
                GOTO Cleanup
            END

            SELECT @UserId = NULL
            SELECT @UserId = UserId FROM dbo.aspnet_Users WHERE LoweredUserName = LOWER(@UserName) AND ApplicationId = @ApplicationId
            IF (@UserId IS NULL)
            BEGIN
                DECLARE @DateTimeNowUTC DATETIME
                EXEC dbo.aspnet_GetUtcDate @TimeZoneAdjustment, @DateTimeNowUTC OUTPUT
                EXEC dbo.aspnet_Users_CreateUser @ApplicationId, @UserName, 0, @DateTimeNowUTC, @UserId OUTPUT
            END

            IF (EXISTS(SELECT * FROM dbo.aspnet_UsersInRoles WHERE @UserId = UserId AND @RoleId = RoleId))
            BEGIN
                SELECT @UserName, @RoleName
                SET @ErrorCode = 3
                GOTO Cleanup
            END
            INSERT INTO dbo.aspnet_UsersInRoles (UserId, RoleId) VALUES(@UserId, @RoleId)
        END
    END

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
        COMMIT TRANSACTION
    END

    RETURN(0)

Cleanup:

    IF( @TranStarted = 1 )
    BEGIN
        SET @TranStarted = 0
        ROLLBACK TRANSACTION
    END

    RETURN @ErrorCode
END

GO

ALTER PROCEDURE [dbo].aspnet_Setup_RestorePermissions
    @name   sysname
AS
BEGIN
    DECLARE @object sysname
    DECLARE @protectType char(10)
    DECLARE @action varchar(20)
    DECLARE @grantee sysname
    DECLARE @cmd nvarchar(500)
    DECLARE c1 CURSOR FORWARD_ONLY FOR
        SELECT Object, ProtectType, [Action], Grantee FROM dbo.temp_aspnet_Permissions where Object = @name

    OPEN c1

    FETCH c1 INTO @object, @protectType, @action, @grantee
    WHILE (@@fetch_status = 0)
    BEGIN
        SET @cmd = @protectType + ' ' + @action + ' on ' + @object + ' TO [' + @grantee + ']'
        EXEC (@cmd)
        FETCH c1 INTO @object, @protectType, @action, @grantee
    END

    close c1
    deallocate c1
END

GO
-- Add primary Key Constraints
ALTER TABLE [dbo].[aspnet_Profile] ADD CONSTRAINT [PK__aspnet_Profile] PRIMARY KEY CLUSTERED  ([UserId])
ALTER TABLE [dbo].[aspnet_Applications] ADD CONSTRAINT [PK__aspnet_Applicatitions] PRIMARY KEY NONCLUSTERED  ([ApplicationId])
ALTER TABLE [dbo].[aspnet_SchemaVersions] ADD CONSTRAINT [PK__aspnet_SchemaVersions] PRIMARY KEY CLUSTERED  ([Feature], [CompatibleSchemaVersion])
ALTER TABLE [dbo].[aspnet_UsersInRoles] ADD CONSTRAINT [PK__aspnet_UsersInRoles] PRIMARY KEY CLUSTERED  ([UserId], [RoleId])
ALTER TABLE [dbo].[aspnet_Roles] ADD CONSTRAINT [PK__aspnet_Roles] PRIMARY KEY NONCLUSTERED  ([RoleId])
ALTER TABLE [dbo].[aspnet_Users] ADD CONSTRAINT [PK__aspnet_Users] PRIMARY KEY NONCLUSTERED  ([UserId])

-- Add Indexes
CREATE UNIQUE CLUSTERED INDEX [aspnet_Users_Index] ON [dbo].[aspnet_Users] ([ApplicationId], [LoweredUserName])
CREATE NONCLUSTERED INDEX [aspnet_Users_Index2] ON [dbo].[aspnet_Users] ([ApplicationId], [LastActivityDate])
CREATE NONCLUSTERED INDEX [aspnet_UsersInRoles_index] ON [dbo].[aspnet_UsersInRoles] ([RoleId])
CREATE UNIQUE CLUSTERED INDEX [aspnet_Roles_index1] ON [dbo].[aspnet_Roles] ([ApplicationId], [LoweredRoleName])

-- Add Unique Constraints
ALTER TABLE [dbo].[aspnet_Applications] ADD CONSTRAINT [UQ__aspnet_Applicatitions_ApplicationName] UNIQUE NONCLUSTERED  ([ApplicationName])
ALTER TABLE [dbo].[aspnet_Applications] ADD CONSTRAINT [UQ__aspnet_Applicatitions_LoweredApplicationName] UNIQUE NONCLUSTERED  ([LoweredApplicationName])

-- Add Default Constraints
ALTER TABLE [dbo].[aspnet_Applications] ADD CONSTRAINT [DF__aspnet_Applications_ApplicationId] DEFAULT (newid()) FOR [ApplicationId]
ALTER TABLE [dbo].[aspnet_Roles] ADD CONSTRAINT [DF__aspnet_Roles_RoleId] DEFAULT (newid()) FOR [RoleId]
ALTER TABLE [dbo].[aspnet_Users] ADD CONSTRAINT [DF__aspnet_Users_UserId] DEFAULT (newid()) FOR [UserId]
ALTER TABLE [dbo].[aspnet_Users] ADD CONSTRAINT [DF__aspnet_Users_MobileAlias] DEFAULT (null) FOR [MobileAlias]
ALTER TABLE [dbo].[aspnet_Users] ADD CONSTRAINT [DF__aspnet_Users_IsAnonymous] DEFAULT (0) FOR [IsAnonymous]

-- Add Foreign Keys
ALTER TABLE [dbo].[aspnet_Membership] ADD
	CONSTRAINT [FK__aspnet_Membership_Applications] FOREIGN KEY ([ApplicationId]) REFERENCES [dbo].[aspnet_Applications] ([ApplicationId]),
	CONSTRAINT [FK__aspnet_Membership_Users] FOREIGN KEY ([UserId]) REFERENCES [dbo].[aspnet_Users] ([UserId])
ALTER TABLE [dbo].[aspnet_Profile] ADD CONSTRAINT [FK__aspnet_Profile_Users] FOREIGN KEY ([UserId]) REFERENCES [dbo].[aspnet_Users] ([UserId])
ALTER TABLE [dbo].[aspnet_Roles] ADD CONSTRAINT [FK__aspnet_Roles_Applications] FOREIGN KEY ([ApplicationId]) REFERENCES [dbo].[aspnet_Applications] ([ApplicationId])
ALTER TABLE [dbo].[aspnet_UsersInRoles] ADD
	CONSTRAINT [FK__aspnet_UsersInRoles_Roles] FOREIGN KEY ([RoleId]) REFERENCES [dbo].[aspnet_Roles] ([RoleId]),
	CONSTRAINT [FK__aspnet_UsersInRoles_Users] FOREIGN KEY ([UserId]) REFERENCES [dbo].[aspnet_Users] ([UserId])
ALTER TABLE [dbo].[aspnet_Applications] ALTER COLUMN [ApplicationId] [uniqueidentifier] NOT NULL
ALTER TABLE [dbo].[aspnet_Users] ADD CONSTRAINT [FK__aspnet_Users_Applications] FOREIGN KEY ([ApplicationId]) REFERENCES [dbo].[aspnet_Applications] ([ApplicationId])
GO

-- Add SQL Server Roles
IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Roles_FullAccess'  ) )
EXEC sp_addrole N'aspnet_Roles_FullAccess'

IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Roles_BasicAccess'  ) )
  BEGIN
	EXEC sp_addrole N'aspnet_Roles_BasicAccess'
	EXEC sp_addrolemember N'aspnet_Roles_BasicAccess', N'aspnet_Roles_FullAccess'
  END

IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Roles_ReportingAccess'  ) )
  BEGIN
	EXEC sp_addrole N'aspnet_Roles_ReportingAccess'
	EXEC sp_addrolemember N'aspnet_Roles_ReportingAccess', N'aspnet_Roles_FullAccess'
  END
GO

IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Profile_FullAccess' ) )
EXEC sp_addrole N'aspnet_Profile_FullAccess'

IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Profile_BasicAccess' ) )
  BEGIN
	EXEC sp_addrole N'aspnet_Profile_BasicAccess'
	EXEC sp_addrolemember N'aspnet_Profile_BasicAccess', N'aspnet_Profile_FullAccess'
  END

IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Profile_ReportingAccess' ) )
  BEGIN
	EXEC sp_addrole N'aspnet_Profile_ReportingAccess'
	EXEC sp_addrolemember N'aspnet_Profile_ReportingAccess', N'aspnet_Profile_FullAccess'
  END
GO

IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Membership_FullAccess'  ) )
EXEC sp_addrole N'aspnet_Membership_FullAccess'

IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Membership_BasicAccess'  ) )
  BEGIN
	EXEC sp_addrole N'aspnet_Membership_BasicAccess'
	EXEC sp_addrolemember N'aspnet_Membership_BasicAccess', N'aspnet_Membership_FullAccess'
  END

IF ( NOT EXISTS ( SELECT name
                  FROM sysusers
                  WHERE issqlrole = 1
                  AND name = N'aspnet_Membership_ReportingAccess'  ) )
  BEGIN
	EXEC sp_addrole N'aspnet_Membership_ReportingAccess'
	EXEC sp_addrolemember N'aspnet_Membership_ReportingAccess', N'aspnet_Membership_FullAccess'
  END
GO

GRANT SELECT ON [dbo].[vw_aspnet_Applications] TO [aspnet_Membership_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_Applications] TO [aspnet_Profile_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_Applications] TO [aspnet_Roles_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_MembershipUsers] TO [aspnet_Membership_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_Profiles] TO [aspnet_Profile_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_Roles] TO [aspnet_Roles_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_Users] TO [aspnet_Membership_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_Users] TO [aspnet_Profile_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_Users] TO [aspnet_Roles_ReportingAccess]
GRANT SELECT ON [dbo].[vw_aspnet_UsersInRoles] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_CheckSchemaVersion] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_CheckSchemaVersion] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_CheckSchemaVersion] TO [aspnet_Profile_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_CheckSchemaVersion] TO [aspnet_Profile_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_CheckSchemaVersion] TO [aspnet_Roles_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_CheckSchemaVersion] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_ChangePasswordQuestionAndAnswer] TO [aspnet_Membership_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_CreateUser] TO [aspnet_Membership_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_FindUsersByEmail] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_FindUsersByName] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetAllUsers] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetNumberOfUsersOnline] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetNumberOfUsersOnline] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetPassword] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetPasswordWithFormat] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetUserByEmail] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetUserByEmail] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetUserByName] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetUserByName] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetUserByUserId] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_GetUserByUserId] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_ResetPassword] TO [aspnet_Membership_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_SetPassword] TO [aspnet_Membership_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_UnlockUser] TO [aspnet_Membership_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_UpdateLastLoginAndActivityDates] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_UpdateUser] TO [aspnet_Membership_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Membership_UpdateUserInfo] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Profile_DeleteInactiveProfiles] TO [aspnet_Profile_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Profile_DeleteProfiles] TO [aspnet_Profile_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Profile_GetNumberOfInactiveProfiles] TO [aspnet_Profile_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Profile_GetProfiles] TO [aspnet_Profile_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Profile_GetProperties] TO [aspnet_Profile_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Profile_SetProperties] TO [aspnet_Profile_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_Roles_CreateRole] TO [aspnet_Roles_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Roles_DeleteRole] TO [aspnet_Roles_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_Roles_GetAllRoles] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Roles_RoleExists] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_Users_DeleteUser] TO [aspnet_Membership_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_UsersInRoles_AddUsersToRoles] TO [aspnet_Roles_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_UsersInRoles_FindUsersInRole] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_UsersInRoles_GetRolesForUser] TO [aspnet_Roles_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_UsersInRoles_GetRolesForUser] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_UsersInRoles_GetUsersInRoles] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_UsersInRoles_IsUserInRole] TO [aspnet_Roles_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_UsersInRoles_IsUserInRole] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_UsersInRoles_RemoveUsersFromRoles] TO [aspnet_Roles_FullAccess]
GRANT EXECUTE ON [dbo].[aspnet_RegisterSchemaVersion] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_RegisterSchemaVersion] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_RegisterSchemaVersion] TO [aspnet_Profile_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_RegisterSchemaVersion] TO [aspnet_Profile_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_RegisterSchemaVersion] TO [aspnet_Roles_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_RegisterSchemaVersion] TO [aspnet_Roles_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_UnRegisterSchemaVersion] TO [aspnet_Membership_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_UnRegisterSchemaVersion] TO [aspnet_Membership_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_UnRegisterSchemaVersion] TO [aspnet_Profile_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_UnRegisterSchemaVersion] TO [aspnet_Profile_ReportingAccess]
GRANT EXECUTE ON [dbo].[aspnet_UnRegisterSchemaVersion] TO [aspnet_Roles_BasicAccess]
GRANT EXECUTE ON [dbo].[aspnet_UnRegisterSchemaVersion] TO [aspnet_Roles_ReportingAccess]
GO


/************************************************************/
/*****              SqlDataProvider                     *****/
/************************************************************/