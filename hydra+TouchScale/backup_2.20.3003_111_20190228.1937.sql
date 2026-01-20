-- MySQL dump 10.14  Distrib 5.5.56-MariaDB, for debian-linux-gnu (i686)
--
-- Host: localhost    Database: pcscale
-- ------------------------------------------------------
-- Server version	5.5.56-MariaDB-1~precise

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `advertising`
--

DROP TABLE IF EXISTS `advertising`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `advertising` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Scope` smallint(6) DEFAULT '0',
  `Number` int(11) DEFAULT '0',
  `Timeout` int(11) DEFAULT '0',
  `ShowTime` int(11) DEFAULT '10',
  `FadeTime` int(11) DEFAULT '0',
  `FadeType` int(11) DEFAULT '0',
  `ZeroOnly` tinyint(1) DEFAULT '1',
  `ForcedImage` tinyint(1) DEFAULT '0',
  `Publi0` smallint(6) DEFAULT '0',
  `Publi1` smallint(6) DEFAULT '0',
  `PlayList0` varchar(32) NOT NULL DEFAULT '',
  `PlayList1` varchar(32) NOT NULL DEFAULT '',
  `FullScreen0` tinyint(1) DEFAULT '0',
  `FullScreen1` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Scope` (`Scope`,`Number`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `advertising`
--

LOCK TABLES `advertising` WRITE;
/*!40000 ALTER TABLE `advertising` DISABLE KEYS */;
INSERT INTO `advertising` VALUES (1,0,0,0,10,0,0,1,0,0,0,'','',0,0);
/*!40000 ALTER TABLE `advertising` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `AdvertisingInserted` AFTER INSERT ON `Advertising`
FOR EACH ROW BEGIN
  IF ((NEW.`Scope`=2) AND (NEW.`Number`=GLOBAL_GET("current_scale"))) OR
     ((NEW.`Scope`=1) AND (NEW.`Number`=GLOBAL_GET("current_counter"))) OR
     ((NEW.`Scope`=0) AND (NEW.`Number`=0)) THEN
    UPDATE `LocalStatus` SET `AdvertisingChanged` = TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `AdvertisingUpdated` AFTER UPDATE ON `Advertising`
FOR EACH ROW BEGIN
  IF ((NEW.`Scope`=2) AND (NEW.`Number`=GLOBAL_GET("current_scale"))) OR
     ((NEW.`Scope`=1) AND (NEW.`Number`=GLOBAL_GET("current_counter"))) OR
     ((NEW.`Scope`=0) AND (NEW.`Number`=0)) THEN
    UPDATE `LocalStatus` SET `AdvertisingChanged` = TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `AdvertisingDeleted` AFTER DELETE ON `Advertising`
FOR EACH ROW BEGIN
  IF ((OLD.`Scope`=2) AND (OLD.`Number`=GLOBAL_GET("current_scale"))) OR
     ((OLD.`Scope`=1) AND (OLD.`Number`=GLOBAL_GET("current_counter"))) THEN
    UPDATE `LocalStatus` SET `AdvertisingChanged` = TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `barcodeean13`
--

DROP TABLE IF EXISTS `barcodeean13`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `barcodeean13` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `NumberType` char(1) DEFAULT NULL,
  `Number` smallint(6) DEFAULT '0',
  `NetworkSales` char(12) NOT NULL,
  `NetworkParked` char(12) NOT NULL DEFAULT '',
  `NetworkLineSLB` char(12) NOT NULL,
  `NetworkTotalSLB` char(12) NOT NULL,
  `NetworkSS` char(12) NOT NULL,
  `NetworkPacket1` char(12) NOT NULL,
  `NetworkPacket2` char(12) NOT NULL,
  `LocalSales` char(12) NOT NULL,
  `LocalLineSLB` char(12) NOT NULL,
  `LocalTotalSLB` char(12) NOT NULL,
  `LocalSS` char(12) NOT NULL,
  `LocalPacket1` char(12) NOT NULL,
  `LocalPacket2` char(12) NOT NULL,
  `LocalTotal` char(12) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `NumberType` (`NumberType`,`Number`)
) ENGINE=InnoDB AUTO_INCREMENT=20 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `barcodeean13`
--

LOCK TABLES `barcodeean13` WRITE;
/*!40000 ALTER TABLE `barcodeean13` DISABLE KEYS */;
INSERT INTO `barcodeean13` VALUES (2,'C',13,'25022NNTTTTT','','26CCCCCiiiii','25022NNTTTTT','20CCCCCiiiii','20CCCCCiiiii','20CCCCCiiiii','25022NNTTTTT','26CCCCCiiiii','25022NNTTTTT','20CCCCCiiiii','20CCCCCiiiii','20CCCCCiiiii','25022NNTTTTT'),(15,'G',0,'25YTTTTIIIII','','','25YTTTTIIIII','26CCCCCiiiii','26CCCCCiiiii','26CCCCCiiiii','2YTTTTTIIIII','20VVViiiiiii','10VVVIIIIIII','20CCCCCiiiii','20CCCCCiiiii','20CCCCCiiiii','2CCCCCCiiiii'),(19,'C',2,'25YTTTTIIIII','','25YTTTTiiiii','25YTTTTIIIII','26CCCCCiiiii','26CCCCCiiiii','26CCCCCiiiii','25YTTTTIIIII','25YTTTTiiiii','25YTTTTIIIII','26CCCCCiiiii','26CCCCCiiiii','26CCCCCiiiii','25YTTTTIIIII');
/*!40000 ALTER TABLE `barcodeean13` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `bovine`
--

DROP TABLE IF EXISTS `bovine`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `bovine` (
  `NumCard` smallint(3) NOT NULL,
  `EarTag` varchar(30) NOT NULL,
  `BirthLand` smallint(6) DEFAULT NULL,
  `BirthDate` date DEFAULT NULL,
  `BreedingLand` smallint(6) DEFAULT NULL,
  `SlaughterLand` smallint(6) DEFAULT NULL,
  `SlaughterReg` varchar(30) DEFAULT NULL,
  `SlaughterDate` date DEFAULT NULL,
  `ButcherLand` smallint(6) DEFAULT NULL,
  `ButcherReg` varchar(30) DEFAULT NULL,
  `Lot` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`NumCard`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `bovine`
--

LOCK TABLES `bovine` WRITE;
/*!40000 ALTER TABLE `bovine` DISABLE KEYS */;
INSERT INTO `bovine` VALUES (1,'123456789/CE',724,'2019-01-01',702,24,'83790856356dfhg','2019-01-21',24,'56745764760','');
/*!40000 ALTER TABLE `bovine` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `cash`
--

DROP TABLE IF EXISTS `cash`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cash` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Scope` smallint(6) DEFAULT '0',
  `Number` int(11) DEFAULT '0',
  `Name` varchar(32) NOT NULL DEFAULT '',
  `IP` int(10) unsigned NOT NULL DEFAULT '0',
  `User` varchar(32) NOT NULL DEFAULT '',
  `Password` varchar(32) NOT NULL DEFAULT '',
  `WaitPaid` char(1) NOT NULL DEFAULT '0',
  `Status` char(1) NOT NULL DEFAULT '0',
  `Oper` int(10) unsigned NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `IP` (`IP`),
  UNIQUE KEY `Scope` (`Scope`,`Number`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `cash`
--

LOCK TABLES `cash` WRITE;
/*!40000 ALTER TABLE `cash` DISABLE KEYS */;
/*!40000 ALTER TABLE `cash` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `CashInserted` AFTER INSERT ON `Cash`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckCash` = TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `CashUpdated` AFTER UPDATE ON `Cash`
FOR EACH ROW BEGIN
  IF ((NEW.`Scope`!=OLD.`Scope`) OR (NEW.`Number`!=OLD.`Number`) OR (NEW.`Status`!=OLD.`Status`)) THEN
    UPDATE `LocalStatus` SET `CheckCash` = TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `CashDeleted` AFTER DELETE ON `Cash`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckCash` = TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `cashdrawer`
--

DROP TABLE IF EXISTS `cashdrawer`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cashdrawer` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Scale` int(11) NOT NULL DEFAULT '0',
  `Vendor` int(11) NOT NULL DEFAULT '0',
  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `Amount` decimal(10,2) NOT NULL DEFAULT '0.00',
  `CashIn` decimal(10,2) NOT NULL DEFAULT '0.00',
  `CashOut` decimal(10,2) NOT NULL DEFAULT '0.00',
  `CashIP` int(10) unsigned NOT NULL DEFAULT '0',
  `Text` varchar(64) NOT NULL DEFAULT '',
  `CashStatus` char(1) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT '-',
  `CashOperId` int(11) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `cashdrawer`
--

LOCK TABLES `cashdrawer` WRITE;
/*!40000 ALTER TABLE `cashdrawer` DISABLE KEYS */;
/*!40000 ALTER TABLE `cashdrawer` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `CashDrawerUpdated` AFTER UPDATE ON `CashDrawer`
FOR EACH ROW BEGIN
  INSERT INTO `TotalPayTickets` (`Date`, `Type`, `NetStat`, `PayMethod`, `Delivered`)
    VALUES (DATE(NEW.`DatTim`), 'T', 'N', 99, (NEW.`CashIn`-NEW.`CashOut`))
    ON DUPLICATE KEY UPDATE `TotalPayTickets`.`Delivered`=`TotalPayTickets`.`Delivered`+VALUES(`Delivered`);
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `cfgsystem`
--

DROP TABLE IF EXISTS `cfgsystem`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cfgsystem` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `ParamKey` varchar(128) NOT NULL DEFAULT '',
  `ParamValue` varchar(128) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `ParamKey` (`ParamKey`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `cfgsystem`
--

LOCK TABLES `cfgsystem` WRITE;
/*!40000 ALTER TABLE `cfgsystem` DISABLE KEYS */;
INSERT INTO `cfgsystem` VALUES (1,'simplified-invoice-serial','1600101');
/*!40000 ALTER TABLE `cfgsystem` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `CfgInserted` AFTER INSERT ON `CfgSystem`
FOR EACH ROW BEGIN
  SET @trx := (SELECT trx_new("BUTTON_UPDATE", "GUI"));
  SET @ret := (SELECT trx_set_param(@trx, "KEYNAME", CONCAT("CFGSYSTEM_", NEW.`ParamKey`)));
  SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", NEW.`ParamValue`));
  SET @ret := (SELECT trx_send(@trx));
  IF(NEW.`ParamKey`="allow-customer-order") THEN
    UPDATE `LocalStatus` SET `CheckCustomerOrders` = TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `CfgUpdated` AFTER UPDATE ON `CfgSystem`
FOR EACH ROW BEGIN
  IF(OLD.`ParamValue`!=NEW.`ParamValue`) THEN
    SET @trx := (SELECT trx_new("BUTTON_UPDATE", "GUI"));
    SET @ret := (SELECT trx_set_param(@trx, "KEYNAME", CONCAT("CFGSYSTEM_", NEW.`ParamKey`)));
    SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", NEW.`ParamValue`));
    SET @ret := (SELECT trx_send(@trx));
    IF((OLD.`ParamKey`="allow-customer-order") OR (NEW.`ParamKey`="allow-customer-order")) THEN
      UPDATE `LocalStatus` SET `CheckCustomerOrders` = TRUE;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `CfgDeleted` AFTER DELETE ON `CfgSystem`
FOR EACH ROW BEGIN
  SET @trx := (SELECT trx_new("BUTTON_UPDATE", "GUI"));
  SET @ret := (SELECT trx_set_param(@trx, "KEYNAME", CONCAT("CFGSYSTEM_", OLD.`ParamKey`)));
  SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", ""));
  SET @ret := (SELECT trx_send(@trx));
  IF(OLD.`ParamKey`="allow-customer-order") THEN
    UPDATE `LocalStatus` SET `CheckCustomerOrders` = TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `clientconfiguration`
--

DROP TABLE IF EXISTS `clientconfiguration`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `clientconfiguration` (
  `TypeNumber` int(11) NOT NULL,
  `Name` varchar(20) NOT NULL DEFAULT '',
  `Icon` varchar(16) NOT NULL DEFAULT 'face',
  `InUse` tinyint(1) DEFAULT '0',
  `IPv4` int(10) unsigned NOT NULL DEFAULT '0',
  `Port` smallint(5) unsigned NOT NULL DEFAULT '1900',
  `ConnectTimeout` tinyint(3) unsigned NOT NULL DEFAULT '2',
  `TckSend` tinyint(1) DEFAULT '1',
  `TckWaitShort` tinyint(3) unsigned NOT NULL DEFAULT '1',
  `TckWaitLong` tinyint(3) unsigned NOT NULL DEFAULT '5',
  `TckTimeout` tinyint(3) unsigned NOT NULL DEFAULT '2',
  `CmdRequest` tinyint(1) DEFAULT '1',
  `CmdWaitShort` tinyint(3) unsigned NOT NULL DEFAULT '1',
  `CmdWaitLong` tinyint(3) unsigned NOT NULL DEFAULT '60',
  `CmdTimeout` tinyint(3) unsigned NOT NULL DEFAULT '2',
  `DateRequest` tinyint(1) DEFAULT '1',
  `DateTimeRequest` int(10) unsigned NOT NULL DEFAULT '3600',
  `UseQMatic` tinyint(1) DEFAULT '0',
  `URL` varchar(255) NOT NULL DEFAULT 'http://demo.q-matic.es:9094',
  `Company` int(11) DEFAULT '0',
  `Shop` int(11) DEFAULT '0',
  PRIMARY KEY (`TypeNumber`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `clientconfiguration`
--

LOCK TABLES `clientconfiguration` WRITE;
/*!40000 ALTER TABLE `clientconfiguration` DISABLE KEYS */;
INSERT INTO `clientconfiguration` VALUES (1,'ECI','face',0,0,1900,2,1,1,5,2,1,1,60,2,1,3600,0,'http://demo.q-matic.es:9094',0,0);
/*!40000 ALTER TABLE `clientconfiguration` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `code128`
--

DROP TABLE IF EXISTS `code128`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `code128` (
  `Counter` int(11) NOT NULL DEFAULT '0',
  `Code1` varchar(32) DEFAULT '',
  `Code2` varchar(32) DEFAULT '',
  `Code3` varchar(32) DEFAULT '',
  `Code4` varchar(32) DEFAULT '',
  `Code5` varchar(32) DEFAULT '',
  `Code6` varchar(32) DEFAULT '',
  `Code7` varchar(32) DEFAULT '',
  `Code8` varchar(32) DEFAULT '',
  `Code9` varchar(32) DEFAULT '',
  `Code10` varchar(32) DEFAULT '',
  PRIMARY KEY (`Counter`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `code128`
--

LOCK TABLES `code128` WRITE;
/*!40000 ALTER TABLE `code128` DISABLE KEYS */;
/*!40000 ALTER TABLE `code128` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `counterfamilies`
--

DROP TABLE IF EXISTS `counterfamilies`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `counterfamilies` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `CounterCode` int(11) NOT NULL,
  `FamilyIndex` int(11) NOT NULL DEFAULT '1',
  `FamilyCode` int(11) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `idcofam` (`CounterCode`,`FamilyIndex`),
  CONSTRAINT `fkcofam` FOREIGN KEY (`CounterCode`) REFERENCES `counters` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=2262 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `counterfamilies`
--

LOCK TABLES `counterfamilies` WRITE;
/*!40000 ALTER TABLE `counterfamilies` DISABLE KEYS */;
INSERT INTO `counterfamilies` VALUES (2227,2,1,201),(2228,2,2,202),(2229,2,3,203),(2230,2,4,204),(2231,2,5,205),(2232,2,6,206),(2233,2,7,207),(2234,2,8,208),(2235,2,9,209),(2236,2,10,210),(2237,3,1,301),(2238,3,2,302),(2239,3,3,310),(2240,3,4,5303),(2241,4,2,501),(2242,4,3,502),(2243,4,4,503),(2244,4,5,505),(2245,5,1,701),(2246,5,2,702),(2247,5,3,703),(2248,5,4,5305),(2249,6,1,601),(2250,6,2,602),(2251,6,3,603),(2252,6,4,604),(2253,6,5,5306),(2254,8,1,2001),(2255,8,2,2002),(2256,8,3,2003),(2257,8,4,2004),(2258,8,5,2005),(2259,8,6,2006),(2260,8,7,2008),(2261,8,8,2009);
/*!40000 ALTER TABLE `counterfamilies` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertCounterFamilies` AFTER INSERT ON `CounterFamilies`
FOR EACH ROW BEGIN
  IF (NEW.`CounterCode`= GLOBAL_GET("current_counter")) THEN
    UPDATE `LocalStatus` SET `CheckCounterFamilies`=TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateCounterFamilies` AFTER UPDATE ON `CounterFamilies`
FOR EACH ROW BEGIN
  IF (OLD.`CounterCode`= GLOBAL_GET("current_counter")) THEN
    UPDATE `LocalStatus` SET `CheckCounterFamilies`=TRUE;
  END IF;
  IF (NEW.`CounterCode`= GLOBAL_GET("current_counter")) THEN
    UPDATE `LocalStatus` SET `CheckCounterFamilies`=TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteCounterFamilies` AFTER DELETE ON `CounterFamilies`
FOR EACH ROW BEGIN
  IF (OLD.`CounterCode`= GLOBAL_GET("current_counter")) THEN
    UPDATE `LocalStatus` SET `CheckCounterFamilies`=TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `countergoals`
--

DROP TABLE IF EXISTS `countergoals`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `countergoals` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `OnDate` date NOT NULL,
  `Counter` int(11) DEFAULT '0',
  `Amount` decimal(10,2) DEFAULT '0.00',
  `AMCustomers` int(11) DEFAULT '0',
  `PMCustomers` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `idx` (`OnDate`,`Counter`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `countergoals`
--

LOCK TABLES `countergoals` WRITE;
/*!40000 ALTER TABLE `countergoals` DISABLE KEYS */;
/*!40000 ALTER TABLE `countergoals` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `counters`
--

DROP TABLE IF EXISTS `counters`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `counters` (
  `Code` int(11) NOT NULL,
  `Name` varchar(64) DEFAULT '',
  `Icon` varchar(127) DEFAULT NULL,
  `TurnProfile` varchar(3) DEFAULT '',
  `TurnProfileGroup` varchar(3) DEFAULT '',
  `TurnIdentifier` char(1) DEFAULT '',
  `Turn` int(10) unsigned DEFAULT '1',
  `TurnsPending` int(10) unsigned NOT NULL DEFAULT '0',
  `TurnDatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `TurnLast` int(10) unsigned NOT NULL DEFAULT '0',
  `TurnLastTime` time NOT NULL DEFAULT '00:00:00',
  `TurnDelay` int(11) NOT NULL DEFAULT '0',
  `TurnOnTime` tinyint(1) NOT NULL DEFAULT '0',
  `TurnMonitor` tinyint(1) DEFAULT '1',
  `Description` varchar(127) DEFAULT '',
  `Text1` varchar(255) DEFAULT '',
  `Text2` varchar(255) DEFAULT '',
  `Text3` varchar(255) DEFAULT '',
  `Text4` varchar(255) DEFAULT '',
  `Text5` varchar(255) DEFAULT '',
  `Text6` varchar(255) DEFAULT '',
  `Text7` varchar(255) DEFAULT '',
  `Text8` varchar(255) DEFAULT '',
  `Text9` varchar(255) DEFAULT '',
  `Text10` varchar(255) DEFAULT '',
  PRIMARY KEY (`Code`),
  UNIQUE KEY `Code` (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `counters`
--

LOCK TABLES `counters` WRITE;
/*!40000 ALTER TABLE `counters` DISABLE KEYS */;
INSERT INTO `counters` VALUES (1,'FRUTAS Y VERDURAS','fruits2 625x348.png','','','',1,0,'2019-02-26 18:46:27',0,'00:00:00',0,0,1,'PLATOS PREPARADOS','','','','','','','','','',''),(2,'CARNICERIA','beef2_640-480.jpg','','','',1,0,'0000-00-00 00:00:00',0,'00:00:00',0,0,1,'CARNICERIA','','','','','','','','','',''),(3,'CHARCUTERIA','sausage 640x427.jpg','','','',1,0,'0000-00-00 00:00:00',0,'00:00:00',0,0,1,'HARCUTERIA','','','','','','','','','',''),(4,'FRUTAS Y VERDURAS','fish 640x480.jpg','','','',2,0,'0000-00-00 00:00:00',0,'00:00:00',0,0,1,'FRUTAS Y VERDURAS','','','','','','','','','',''),(5,'PESCADERIA','ensalada_verano.png','','','',1,0,'0000-00-00 00:00:00',0,'00:00:00',0,0,1,'PESCADERIA','','','','','','','','','',''),(6,'PANADERIA','bread 640x480.jpg','','','',1,0,'0000-00-00 00:00:00',0,'00:00:00',0,0,1,'PANADERIA','','','','','','','','','',''),(8,'COCINA',NULL,'','','',1,0,'0000-00-00 00:00:00',0,'00:00:00',0,0,1,'COCINA','','','','','','','','','','');
/*!40000 ALTER TABLE `counters` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertCounter` AFTER INSERT ON `Counters`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `CheckCounters`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateCounter` AFTER UPDATE ON `Counters`
FOR EACH ROW BEGIN
  IF(NEW.`Code`!=OLD.`Code`) THEN
    UPDATE `LocalStatus` SET `CheckCounters`=TRUE;
  ELSE
    UPDATE `LocalStatus` SET `CheckCountersKeys`=TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteCounter` AFTER DELETE ON `Counters`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `CheckCounters`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `countries`
--

DROP TABLE IF EXISTS `countries`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `countries` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Locale` char(2) NOT NULL DEFAULT '',
  `Code` int(11) NOT NULL DEFAULT '0',
  `Alpha2` char(2) NOT NULL DEFAULT '',
  `Name` varchar(127) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Locale` (`Locale`,`Code`),
  KEY `Name` (`Name`(19))
) ENGINE=InnoDB AUTO_INCREMENT=748 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `countries`
--

LOCK TABLES `countries` WRITE;
/*!40000 ALTER TABLE `countries` DISABLE KEYS */;
INSERT INTO `countries` VALUES (1,'es',4,'AF','Afganistán'),(2,'es',248,'AX','Islas Äland'),(3,'es',8,'AL','Albania'),(4,'es',12,'DZ','Algeria'),(5,'es',16,'AS','Samoa Americana'),(6,'es',20,'AD','Andorra'),(7,'es',24,'AO','Angola'),(8,'es',660,'AI','Anguila'),(9,'es',10,'AQ','Antártida'),(10,'es',28,'AG','Antigua y Barbuda'),(11,'es',32,'AR','Argentina'),(12,'es',51,'AM','Armenia'),(13,'es',533,'AW','Aruba'),(14,'es',36,'AU','Australia'),(15,'es',40,'AT','Austria'),(16,'es',31,'AZ','Azerbayán'),(17,'es',44,'BS','Bahamas'),(18,'es',48,'BH','Bahrein'),(19,'es',50,'BD','Bangladesh'),(20,'es',52,'BB','Barbados'),(21,'es',112,'BY','Bielorrusia'),(22,'es',56,'BE','Bélgica'),(23,'es',84,'BZ','Belice'),(24,'es',204,'BJ','Benín'),(25,'es',60,'BM','Islas Bermudas'),(26,'es',64,'BT','Bhután'),(27,'es',68,'BO','Bolivia, Estado plurinacional de'),(28,'es',535,'BQ','Islas BES (Caribe Neerlandés)'),(29,'es',70,'BA','Bosnia y Herzegovina'),(30,'es',72,'BW','Botsuana'),(31,'es',74,'BV','Isla Bouvet'),(32,'es',76,'BR','Brasil'),(33,'es',86,'IO','Británico del Océano Índico, Territorio'),(34,'es',96,'BN','Brunei Darussalam'),(35,'es',100,'BG','Bulgaria'),(36,'es',854,'BF','Burquina Faso'),(37,'es',108,'BI','Burundi'),(38,'es',116,'KH','Camboya'),(39,'es',120,'CM','Camerún'),(40,'es',124,'CA','Canadá'),(41,'es',132,'CV','Cabo Verde'),(42,'es',136,'KY','Islas Caimán'),(43,'es',140,'CF','Centro-africana, República'),(44,'es',148,'TD','Chad'),(45,'es',152,'CL','Chile'),(46,'es',156,'CN','China'),(47,'es',162,'CX','Isla de Navidad'),(48,'es',166,'CC','Islas Cocos (Keeling)'),(49,'es',170,'CO','Colombia'),(50,'es',174,'KM','Comores, Islas'),(51,'es',178,'CG','Congo'),(52,'es',180,'CD','Congo, República Democrática del'),(53,'es',184,'CK','Islas Cook'),(54,'es',188,'CR','Costa Rica'),(55,'es',384,'CI','Costa de Marfíl'),(56,'es',191,'HR','Croacia'),(57,'es',192,'CU','Cuba'),(58,'es',531,'CW','Curasao'),(59,'es',196,'CY','Chipre'),(60,'es',203,'CZ','República Checa'),(61,'es',208,'DK','Dinamarca'),(62,'es',262,'DJ','Yibuti'),(63,'es',212,'DM','Dominica'),(64,'es',214,'DO','República Dominicana'),(65,'es',218,'EC','Ecuador'),(66,'es',818,'EG','Egipto'),(67,'es',222,'SV','El Salvador'),(68,'es',226,'GQ','Guinea Ecuatorial'),(69,'es',232,'ER','Eritrea'),(70,'es',233,'EE','Estonia'),(71,'es',231,'ET','Etiopía'),(72,'es',238,'FK','Islas Falkland (Malvinas)'),(73,'es',234,'FO','Islas Feroe'),(74,'es',242,'FJ','Fiyi'),(75,'es',246,'FI','Finlandia'),(76,'es',250,'FR','Francia'),(77,'es',254,'GF','Guayana Francesa'),(78,'es',258,'PF','Polinesia Francesa'),(79,'es',260,'TF','Territorios Franceses del Sur'),(80,'es',266,'GA','Gabón'),(81,'es',270,'GM','Gambia'),(82,'es',268,'GE','Georgia'),(83,'es',276,'DE','Alemania'),(84,'es',288,'GH','Ghana'),(85,'es',292,'GI','Gibraltar'),(86,'es',300,'GR','Grecia'),(87,'es',304,'GL','Groenlandia'),(88,'es',308,'GD','Granada'),(89,'es',312,'GP','Guadalupe'),(90,'es',316,'GU','Guam'),(91,'es',320,'GT','Guatemala'),(92,'es',831,'GG','Guernsey'),(93,'es',324,'GN','Guinea'),(94,'es',624,'GW','Guinea-Bissau'),(95,'es',328,'GY','Guyana'),(96,'es',332,'HT','Haití'),(97,'es',334,'HM','Islas Heard, Islas y McDonald'),(98,'es',336,'VA','Santa Sede (Ciudad Estado del Vaticano)'),(99,'es',340,'HN','Honduras'),(100,'es',344,'HK','Hong Kong'),(101,'es',348,'HU','Hungría'),(102,'es',352,'IS','Islandia'),(103,'es',356,'IN','India'),(104,'es',360,'ID','Indonesia'),(105,'es',364,'IR','Irán, República islámica de'),(106,'es',368,'IQ','Irak'),(107,'es',372,'IE','Irlanda'),(108,'es',833,'IM','Isla de Man'),(109,'es',376,'IL','Israel'),(110,'es',380,'IT','Italia'),(111,'es',388,'JM','Jamaica'),(112,'es',392,'JP','Japón'),(113,'es',832,'JE','Jersey'),(114,'es',400,'JO','Jordania'),(115,'es',398,'KZ','Kazajistán'),(116,'es',404,'KE','Kenia'),(117,'es',296,'KI','Kiribati'),(118,'es',408,'KP','Corea, República Democrática Popular de'),(119,'es',410,'KR','Corea, República de'),(120,'es',414,'KW','Kuwait'),(121,'es',417,'KG','Kirgizstán'),(122,'es',418,'LA','República Democrática Popular de Lao'),(123,'es',428,'LV','Letonia'),(124,'es',422,'LB','Líbano'),(125,'es',426,'LS','Lesoto'),(126,'es',430,'LR','Liberia'),(127,'es',434,'LY','Libia'),(128,'es',438,'LI','Liechtenstein'),(129,'es',440,'LT','Lituania'),(130,'es',442,'LU','Luxemburgo'),(131,'es',446,'MO','Macao'),(132,'es',807,'MK','Macedonia, República de'),(133,'es',450,'MG','Madagascar'),(134,'es',454,'MW','Malawi'),(135,'es',458,'MY','Malasia'),(136,'es',462,'MV','Islas Maldivas'),(137,'es',466,'ML','Mali'),(138,'es',470,'MT','Malta'),(139,'es',584,'MH','Islas Marshall'),(140,'es',474,'MQ','Martinica'),(141,'es',478,'MR','Mauritania'),(142,'es',480,'MU','Mauricio'),(143,'es',175,'YT','Mayotte'),(144,'es',484,'MX','México'),(145,'es',583,'FM','Micronesia, Estados Federados de'),(146,'es',498,'MD','Moldavia, República de'),(147,'es',492,'MC','Mónaco'),(148,'es',496,'MN','Mongolia'),(149,'es',499,'ME','Montenegro'),(150,'es',500,'MS','Montserrat'),(151,'es',504,'MA','Marruecos'),(152,'es',508,'MZ','Mozambique'),(153,'es',104,'MM','Birmania'),(154,'es',516,'NA','Namibia'),(155,'es',520,'NR','Nauru'),(156,'es',524,'NP','Nepal'),(157,'es',528,'NL','Países Bajos'),(158,'es',540,'NC','Nueva Caledonia'),(159,'es',554,'NZ','Nueva Zelanda'),(160,'es',558,'NI','Nicaragua'),(161,'es',562,'NE','Niger'),(162,'es',566,'NG','Nigeria'),(163,'es',570,'NU','Niue'),(164,'es',574,'NF','Isla Norfolk'),(165,'es',580,'MP','Islas Marianas del Norte'),(166,'es',578,'NO','Noruega'),(167,'es',512,'OM','Omán'),(168,'es',586,'PK','Pakistán'),(169,'es',585,'PW','Palau'),(170,'es',275,'PS','Territorio palestino, Ocupado'),(171,'es',591,'PA','Panamá'),(172,'es',598,'PG','Papúa Nueva Guinea'),(173,'es',600,'PY','Paraguay'),(174,'es',604,'PE','Perú'),(175,'es',608,'PH','Filipinas'),(176,'es',612,'PN','Pitcairn'),(177,'es',616,'PL','Polonia'),(178,'es',620,'PT','Portugal'),(179,'es',630,'PR','Puerto Rico'),(180,'es',634,'QA','Qatar'),(181,'es',638,'RE','Reunión'),(182,'es',642,'RO','Rumanía'),(183,'es',643,'RU','Federación Rusa'),(184,'es',646,'RW','Ruanda'),(185,'es',652,'BL','San Bartolomé'),(186,'es',654,'SH','Santa Elena, Ascensión y Tristán de Acuña'),(187,'es',659,'KN','San Cristobo y Nevis'),(188,'es',662,'LC','Santa Lucía'),(189,'es',663,'MF','San Martín (zona francesa)'),(190,'es',666,'PM','San Pedro y Miquelon'),(191,'es',670,'VC','San Vicente y las Granadinas'),(192,'es',882,'WS','Samoa'),(193,'es',674,'SM','San Marino'),(194,'es',678,'ST','Santo Tomé y Príncipe'),(195,'es',682,'SA','Arabia Saudí'),(196,'es',686,'SN','Senegal'),(197,'es',688,'RS','Serbia'),(198,'es',690,'SC','Seychelles'),(199,'es',694,'SL','Sierra Leona'),(200,'es',702,'SG','Singapur'),(201,'es',534,'SX','Isla de San Martín (zona holandsea)'),(202,'es',703,'SK','Eslovaquia'),(203,'es',705,'SI','Eslovenia'),(204,'es',90,'SB','Islas Salomón'),(205,'es',706,'SO','Somalia'),(206,'es',710,'ZA','Suráfrica'),(207,'es',239,'GS','Georgia del Sur e Islas Sandwitch del Sur'),(208,'es',724,'ES','España'),(209,'es',144,'LK','Sri Lanka'),(210,'es',729,'SD','Sudán'),(211,'es',740,'SR','Surinám'),(212,'es',728,'SS','Sudán del Sur'),(213,'es',744,'SJ','Svalbard y Jan Mayen'),(214,'es',748,'SZ','Swazilandia'),(215,'es',752,'SE','Suecia'),(216,'es',756,'CH','Suiza'),(217,'es',760,'SY','República árabe de Siria'),(218,'es',158,'TW','Taiwán, Provincia de China'),(219,'es',762,'TJ','Tadjikistán'),(220,'es',834,'TZ','Tanzania, República unida de'),(221,'es',764,'TH','Tailandia'),(222,'es',626,'TL','Timor-Leste'),(223,'es',768,'TG','Togo'),(224,'es',772,'TK','Tokelau'),(225,'es',776,'TO','Tonga'),(226,'es',780,'TT','Trinidad y Tobago'),(227,'es',788,'TN','Tunez'),(228,'es',792,'TR','Turquía'),(229,'es',795,'TM','Turkmenistán'),(230,'es',796,'TC','Turks y Caicos, Islas'),(231,'es',798,'TV','Tuvalu'),(232,'es',800,'UG','Uganda'),(233,'es',804,'UA','Ucrania'),(234,'es',784,'AE','Emiratos Árabes Unidos'),(235,'es',826,'GB','Reino Unido'),(236,'es',840,'US','Estados Unidos'),(237,'es',581,'UM','Islas menores exteriores de Estados Unidos'),(238,'es',858,'UY','Uruguay'),(239,'es',860,'UZ','Uzbekistán'),(240,'es',548,'VU','Vanuatu'),(241,'es',862,'VE','Venezuela, República Bolivariana de'),(242,'es',704,'VN','Vietnam'),(243,'es',92,'VG','Islas Vírgenes, Británicas'),(244,'es',850,'VI','Islas Vírgenes, de EEUU'),(245,'es',876,'WF','Wallis y Futuna'),(246,'es',732,'EH','Sahara Occidental'),(247,'es',887,'YE','Yemen'),(248,'es',894,'ZM','Zambia'),(249,'es',716,'ZW','Zimbabue'),(250,'en',4,'AF','Afghanistan'),(251,'en',248,'AX','Åland Islands'),(252,'en',8,'AL','Albania'),(253,'en',12,'DZ','Algeria'),(254,'en',16,'AS','American Samoa'),(255,'en',20,'AD','Andorra'),(256,'en',24,'AO','Angola'),(257,'en',660,'AI','Anguilla'),(258,'en',10,'AQ','Antarctica'),(259,'en',28,'AG','Antigua and Barbuda'),(260,'en',32,'AR','Argentina'),(261,'en',51,'AM','Armenia'),(262,'en',533,'AW','Aruba'),(263,'en',36,'AU','Australia'),(264,'en',40,'AT','Austria'),(265,'en',31,'AZ','Azerbaijan'),(266,'en',44,'BS','Bahamas'),(267,'en',48,'BH','Bahrain'),(268,'en',50,'BD','Bangladesh'),(269,'en',52,'BB','Barbados'),(270,'en',112,'BY','Belarus'),(271,'en',56,'BE','Belgium'),(272,'en',84,'BZ','Belize'),(273,'en',204,'BJ','Benin'),(274,'en',60,'BM','Bermuda'),(275,'en',64,'BT','Bhutan'),(276,'en',68,'BO','Bolivia, Plurinational State of'),(277,'en',535,'BQ','Bonaire, Sint Eustatius and Saba'),(278,'en',70,'BA','Bosnia and Herzegovina'),(279,'en',72,'BW','Botswana'),(280,'en',74,'BV','Bouvet Island'),(281,'en',76,'BR','Brazil'),(282,'en',86,'IO','British Indian Ocean Territory'),(283,'en',96,'BN','Brunei Darussalam'),(284,'en',100,'BG','Bulgaria'),(285,'en',854,'BF','Burkina Faso'),(286,'en',108,'BI','Burundi'),(287,'en',116,'KH','Cambodia'),(288,'en',120,'CM','Cameroon'),(289,'en',124,'CA','Canada'),(290,'en',132,'CV','Cape Verde'),(291,'en',136,'KY','Cayman Islands'),(292,'en',140,'CF','Central African Republic'),(293,'en',148,'TD','Chad'),(294,'en',152,'CL','Chile'),(295,'en',156,'CN','China'),(296,'en',162,'CX','Christmas Island'),(297,'en',166,'CC','Cocos (Keeling) Islands'),(298,'en',170,'CO','Colombia'),(299,'en',174,'KM','Comoros'),(300,'en',178,'CG','Congo'),(301,'en',180,'CD','Congo, The Democratic Republic of the'),(302,'en',184,'CK','Cook Islands'),(303,'en',188,'CR','Costa Rica'),(304,'en',384,'CI','Côte d\'Ivoire'),(305,'en',191,'HR','Croatia'),(306,'en',192,'CU','Cuba'),(307,'en',531,'CW','Curaçao'),(308,'en',196,'CY','Cyprus'),(309,'en',203,'CZ','Czech Republic'),(310,'en',208,'DK','Denmark'),(311,'en',262,'DJ','Djibouti'),(312,'en',212,'DM','Dominica'),(313,'en',214,'DO','Dominican Republic'),(314,'en',218,'EC','Ecuador'),(315,'en',818,'EG','Egypt'),(316,'en',222,'SV','El Salvador'),(317,'en',226,'GQ','Equatorial Guinea'),(318,'en',232,'ER','Eritrea'),(319,'en',233,'EE','Estonia'),(320,'en',231,'ET','Ethiopia'),(321,'en',238,'FK','Falkland Islands (Malvinas)'),(322,'en',234,'FO','Faroe Islands'),(323,'en',242,'FJ','Fiji'),(324,'en',246,'FI','Finland'),(325,'en',250,'FR','France'),(326,'en',254,'GF','French Guiana'),(327,'en',258,'PF','French Polynesia'),(328,'en',260,'TF','French Southern Territories'),(329,'en',266,'GA','Gabon'),(330,'en',270,'GM','Gambia'),(331,'en',268,'GE','Georgia'),(332,'en',276,'DE','Germany'),(333,'en',288,'GH','Ghana'),(334,'en',292,'GI','Gibraltar'),(335,'en',300,'GR','Greece'),(336,'en',304,'GL','Greenland'),(337,'en',308,'GD','Grenada'),(338,'en',312,'GP','Guadeloupe'),(339,'en',316,'GU','Guam'),(340,'en',320,'GT','Guatemala'),(341,'en',831,'GG','Guernsey'),(342,'en',324,'GN','Guinea'),(343,'en',624,'GW','Guinea-Bissau'),(344,'en',328,'GY','Guyana'),(345,'en',332,'HT','Haiti'),(346,'en',334,'HM','Heard Island and McDonald Islands'),(347,'en',336,'VA','Holy See (Vatican City State)'),(348,'en',340,'HN','Honduras'),(349,'en',344,'HK','Hong Kong'),(350,'en',348,'HU','Hungary'),(351,'en',352,'IS','Iceland'),(352,'en',356,'IN','India'),(353,'en',360,'ID','Indonesia'),(354,'en',364,'IR','Iran, Islamic Republic of'),(355,'en',368,'IQ','Iraq'),(356,'en',372,'IE','Ireland'),(357,'en',833,'IM','Isle of Man'),(358,'en',376,'IL','Israel'),(359,'en',380,'IT','Italy'),(360,'en',388,'JM','Jamaica'),(361,'en',392,'JP','Japan'),(362,'en',832,'JE','Jersey'),(363,'en',400,'JO','Jordan'),(364,'en',398,'KZ','Kazakhstan'),(365,'en',404,'KE','Kenya'),(366,'en',296,'KI','Kiribati'),(367,'en',408,'KP','Korea, Democratic People\'s Republic of'),(368,'en',410,'KR','Korea, Republic of'),(369,'en',414,'KW','Kuwait'),(370,'en',417,'KG','Kyrgyzstan'),(371,'en',418,'LA','Lao People\'s Democratic Republic'),(372,'en',428,'LV','Latvia'),(373,'en',422,'LB','Lebanon'),(374,'en',426,'LS','Lesotho'),(375,'en',430,'LR','Liberia'),(376,'en',434,'LY','Libya'),(377,'en',438,'LI','Liechtenstein'),(378,'en',440,'LT','Lithuania'),(379,'en',442,'LU','Luxembourg'),(380,'en',446,'MO','Macao'),(381,'en',807,'MK','Macedonia, Republic of'),(382,'en',450,'MG','Madagascar'),(383,'en',454,'MW','Malawi'),(384,'en',458,'MY','Malaysia'),(385,'en',462,'MV','Maldives'),(386,'en',466,'ML','Mali'),(387,'en',470,'MT','Malta'),(388,'en',584,'MH','Marshall Islands'),(389,'en',474,'MQ','Martinique'),(390,'en',478,'MR','Mauritania'),(391,'en',480,'MU','Mauritius'),(392,'en',175,'YT','Mayotte'),(393,'en',484,'MX','Mexico'),(394,'en',583,'FM','Micronesia, Federated States of'),(395,'en',498,'MD','Moldova, Republic of'),(396,'en',492,'MC','Monaco'),(397,'en',496,'MN','Mongolia'),(398,'en',499,'ME','Montenegro'),(399,'en',500,'MS','Montserrat'),(400,'en',504,'MA','Morocco'),(401,'en',508,'MZ','Mozambique'),(402,'en',104,'MM','Myanmar'),(403,'en',516,'NA','Namibia'),(404,'en',520,'NR','Nauru'),(405,'en',524,'NP','Nepal'),(406,'en',528,'NL','Netherlands'),(407,'en',540,'NC','New Caledonia'),(408,'en',554,'NZ','New Zealand'),(409,'en',558,'NI','Nicaragua'),(410,'en',562,'NE','Niger'),(411,'en',566,'NG','Nigeria'),(412,'en',570,'NU','Niue'),(413,'en',574,'NF','Norfolk Island'),(414,'en',580,'MP','Northern Mariana Islands'),(415,'en',578,'NO','Norway'),(416,'en',512,'OM','Oman'),(417,'en',586,'PK','Pakistan'),(418,'en',585,'PW','Palau'),(419,'en',275,'PS','Palestinian Territory, Occupied'),(420,'en',591,'PA','Panama'),(421,'en',598,'PG','Papua New Guinea'),(422,'en',600,'PY','Paraguay'),(423,'en',604,'PE','Peru'),(424,'en',608,'PH','Philippines'),(425,'en',612,'PN','Pitcairn'),(426,'en',616,'PL','Poland'),(427,'en',620,'PT','Portugal'),(428,'en',630,'PR','Puerto Rico'),(429,'en',634,'QA','Qatar'),(430,'en',638,'RE','Reunion'),(431,'en',642,'RO','Romania'),(432,'en',643,'RU','Russian Federation'),(433,'en',646,'RW','Rwanda'),(434,'en',652,'BL','Saint Barthélemy'),(435,'en',654,'SH','Saint Helena, Ascension and Tristan da Cunha'),(436,'en',659,'KN','Saint Kitts and Nevis'),(437,'en',662,'LC','Saint Lucia'),(438,'en',663,'MF','Saint Martin (French part)'),(439,'en',666,'PM','Saint Pierre and Miquelon'),(440,'en',670,'VC','Saint Vincent and the Grenadines'),(441,'en',882,'WS','Samoa'),(442,'en',674,'SM','San Marino'),(443,'en',678,'ST','Sao Tome and Principe'),(444,'en',682,'SA','Saudi Arabia'),(445,'en',686,'SN','Senegal'),(446,'en',688,'RS','Serbia'),(447,'en',690,'SC','Seychelles'),(448,'en',694,'SL','Sierra Leone'),(449,'en',702,'SG','Singapore'),(450,'en',534,'SX','Sint Maarten (Dutch part)'),(451,'en',703,'SK','Slovakia'),(452,'en',705,'SI','Slovenia'),(453,'en',90,'SB','Solomon Islands'),(454,'en',706,'SO','Somalia'),(455,'en',710,'ZA','South Africa'),(456,'en',239,'GS','South Georgia and the South Sandwich Islands'),(457,'en',724,'ES','Spain'),(458,'en',144,'LK','Sri Lanka'),(459,'en',729,'SD','Sudan'),(460,'en',740,'SR','Suriname'),(461,'en',728,'SS','South Sudan'),(462,'en',744,'SJ','Svalbard and Jan Mayen'),(463,'en',748,'SZ','Swaziland'),(464,'en',752,'SE','Sweden'),(465,'en',756,'CH','Switzerland'),(466,'en',760,'SY','Syrian Arab Republic'),(467,'en',158,'TW','Taiwan, Province of China'),(468,'en',762,'TJ','Tajikistan'),(469,'en',834,'TZ','Tanzania, United Republic of'),(470,'en',764,'TH','Thailand'),(471,'en',626,'TL','Timor-Leste'),(472,'en',768,'TG','Togo'),(473,'en',772,'TK','Tokelau'),(474,'en',776,'TO','Tonga'),(475,'en',780,'TT','Trinidad and Tobago'),(476,'en',788,'TN','Tunisia'),(477,'en',792,'TR','Turkey'),(478,'en',795,'TM','Turkmenistan'),(479,'en',796,'TC','Turks and Caicos Islands'),(480,'en',798,'TV','Tuvalu'),(481,'en',800,'UG','Uganda'),(482,'en',804,'UA','Ukraine'),(483,'en',784,'AE','United Arab Emirates'),(484,'en',826,'GB','United Kingdom'),(485,'en',840,'US','United States'),(486,'en',581,'UM','United States Minor Outlying Islands'),(487,'en',858,'UY','Uruguay'),(488,'en',860,'UZ','Uzbekistan'),(489,'en',548,'VU','Vanuatu'),(490,'en',862,'VE','Venezuela, Bolivarian Republic of'),(491,'en',704,'VN','Viet Nam'),(492,'en',92,'VG','Virgin Islands, British'),(493,'en',850,'VI','Virgin Islands, U.S.'),(494,'en',876,'WF','Wallis and Futuna'),(495,'en',732,'EH','Western Sahara'),(496,'en',887,'YE','Yemen'),(497,'en',894,'ZM','Zambia'),(498,'en',716,'ZW','Zimbabwe'),(499,'ca',4,'AF','Afganistan'),(500,'ca',248,'AX','Illes Aland'),(501,'ca',8,'AL','Albània'),(502,'ca',12,'DZ','Algèria'),(503,'ca',16,'AS','Samoa Americana'),(504,'ca',20,'AD','Andorra'),(505,'ca',24,'AO','Angola'),(506,'ca',660,'AI','Anguilla'),(507,'ca',10,'AQ','Antàrtida'),(508,'ca',28,'AG','Antigua i Barbuda'),(509,'ca',32,'AR','Argentina'),(510,'ca',51,'AM','Armènia'),(511,'ca',533,'AW','Aruba'),(512,'ca',36,'AU','Austràlia'),(513,'ca',40,'AT','Àustria'),(514,'ca',31,'AZ','Azerbaitjan'),(515,'ca',44,'BS','Bahames'),(516,'ca',48,'BH','Bahrain'),(517,'ca',50,'BD','Bangla Desh'),(518,'ca',52,'BB','Barbados'),(519,'ca',112,'BY','Bielorússia'),(520,'ca',56,'BE','Bèlgica'),(521,'ca',84,'BZ','Belize'),(522,'ca',204,'BJ','Benín'),(523,'ca',60,'BM','Bermuda'),(524,'ca',64,'BT','Bhutan'),(525,'ca',68,'BO','Bolívia, Estat Plurinacional de'),(526,'ca',535,'BQ','Bonaire, Sint Eustatius and Saba'),(527,'ca',70,'BA','Bòsnia i Hercegovina'),(528,'ca',72,'BW','Botswana'),(529,'ca',74,'BV','Illa Bouvet'),(530,'ca',76,'BR','Brasil'),(531,'ca',86,'IO','Territori Britànic de l\'Oceà Índic'),(532,'ca',96,'BN','Brunei (Negara Brunei Darussalam)'),(533,'ca',100,'BG','Bulgària'),(534,'ca',854,'BF','Burkina Faso'),(535,'ca',108,'BI','Burundi'),(536,'ca',116,'KH','Cambodja'),(537,'ca',120,'CM','Camerun'),(538,'ca',124,'CA','Canadà'),(539,'ca',132,'CV','Cap Verd'),(540,'ca',136,'KY','Illes Caiman'),(541,'ca',140,'CF','República Centreafricana'),(542,'ca',148,'TD','Txad'),(543,'ca',152,'CL','Xile'),(544,'ca',156,'CN','Xina'),(545,'ca',162,'CX','Illa Christmas'),(546,'ca',166,'CC','Illes Cocos (Keeling)'),(547,'ca',170,'CO','Colòmbia'),(548,'ca',174,'KM','Comores'),(549,'ca',178,'CG','Congo'),(550,'ca',180,'CD','Congo, La República Democràtica del'),(551,'ca',184,'CK','Illes Cook'),(552,'ca',188,'CR','Costa Rica'),(553,'ca',384,'CI','Costa de Vori'),(554,'ca',191,'HR','Croàcia'),(555,'ca',192,'CU','Cuba'),(556,'ca',531,'CW','Curaçao'),(557,'ca',196,'CY','Xipre'),(558,'ca',203,'CZ','República Txeca'),(559,'ca',208,'DK','Dinamarca'),(560,'ca',262,'DJ','Djibouti'),(561,'ca',212,'DM','Dominica'),(562,'ca',214,'DO','República Dominicana'),(563,'ca',218,'EC','Equador'),(564,'ca',818,'EG','Egipte'),(565,'ca',222,'SV','El Salvador'),(566,'ca',226,'GQ','Guinea Equatorial'),(567,'ca',232,'ER','Eritrea'),(568,'ca',233,'EE','Estònia'),(569,'ca',231,'ET','Etiòpia'),(570,'ca',238,'FK','Illes Malvines (Falkland)'),(571,'ca',234,'FO','Illes Fèroe'),(572,'ca',242,'FJ','Fiji'),(573,'ca',246,'FI','Finlàndia'),(574,'ca',250,'FR','França'),(575,'ca',254,'GF','Guaiana Francesa'),(576,'ca',258,'PF','Polinèsia francesa'),(577,'ca',260,'TF','Territoris Francesos del Sud'),(578,'ca',266,'GA','Gabon'),(579,'ca',270,'GM','Gàmbia'),(580,'ca',268,'GE','Geòrgia'),(581,'ca',276,'DE','Alemanya'),(582,'ca',288,'GH','Ghana'),(583,'ca',292,'GI','Gibraltar'),(584,'ca',300,'GR','Grècia'),(585,'ca',304,'GL','Groenlàndia'),(586,'ca',308,'GD','Grenada'),(587,'ca',312,'GP','Guadalupe'),(588,'ca',316,'GU','Guam'),(589,'ca',320,'GT','Guatemala'),(590,'ca',831,'GG','Guernsey'),(591,'ca',324,'GN','Guinea'),(592,'ca',624,'GW','Guinea Bissau'),(593,'ca',328,'GY','Guyana'),(594,'ca',332,'HT','Haití'),(595,'ca',334,'HM','Illa Heard i Illes McDonald'),(596,'ca',336,'VA','Santa Seu (Estat del Vaticà)'),(597,'ca',340,'HN','Hondures'),(598,'ca',344,'HK','Hong Kong'),(599,'ca',348,'HU','Hongria'),(600,'ca',352,'IS','Islàndia'),(601,'ca',356,'IN','Índia'),(602,'ca',360,'ID','Indonèsia'),(603,'ca',364,'IR','Iran'),(604,'ca',368,'IQ','Iraq'),(605,'ca',372,'IE','Irlanda'),(606,'ca',833,'IM','Illa de Man'),(607,'ca',376,'IL','Israel'),(608,'ca',380,'IT','Itàlia'),(609,'ca',388,'JM','Jamaica'),(610,'ca',392,'JP','Japó'),(611,'ca',832,'JE','Jersey'),(612,'ca',400,'JO','Jordània'),(613,'ca',398,'KZ','Kazakhstan'),(614,'ca',404,'KE','Kenya'),(615,'ca',296,'KI','Kiribati'),(616,'ca',408,'KP','Corea del Nord'),(617,'ca',410,'KR','Corea del Sud'),(618,'ca',414,'KW','Kuwait'),(619,'ca',417,'KG','Kirguizistan'),(620,'ca',418,'LA','República Democràtica Popular de Laos'),(621,'ca',428,'LV','Letònia'),(622,'ca',422,'LB','Líban'),(623,'ca',426,'LS','Lesotho'),(624,'ca',430,'LR','Libèria'),(625,'ca',434,'LY','Líbia'),(626,'ca',438,'LI','Liechtenstein'),(627,'ca',440,'LT','Lituània'),(628,'ca',442,'LU','Luxemburg'),(629,'ca',446,'MO','Macau'),(630,'ca',807,'MK','Macedònia, República de'),(631,'ca',450,'MG','Madagascar'),(632,'ca',454,'MW','Malawi'),(633,'ca',458,'MY','Malàisia'),(634,'ca',462,'MV','Maldives'),(635,'ca',466,'ML','Mali'),(636,'ca',470,'MT','Malta'),(637,'ca',584,'MH','Illes Marshall'),(638,'ca',474,'MQ','Martinica'),(639,'ca',478,'MR','Mauritània'),(640,'ca',480,'MU','Maurici'),(641,'ca',175,'YT','Mayotte'),(642,'ca',484,'MX','Mèxic'),(643,'ca',583,'FM','Micronèsia, Estats Federats de'),(644,'ca',498,'MD','Moldàvia, República de'),(645,'ca',492,'MC','Mònaco'),(646,'ca',496,'MN','Mongòlia'),(647,'ca',499,'ME','Montenegro'),(648,'ca',500,'MS','Montserrat'),(649,'ca',504,'MA','Marroc'),(650,'ca',508,'MZ','Moçambic'),(651,'ca',104,'MM','Myanmar'),(652,'ca',516,'NA','Namíbia'),(653,'ca',520,'NR','Nauru'),(654,'ca',524,'NP','Nepal'),(655,'ca',528,'NL','Països Baixos'),(656,'ca',540,'NC','Nova Caledònia'),(657,'ca',554,'NZ','Nova Zelanda'),(658,'ca',558,'NI','Nicaragua'),(659,'ca',562,'NE','Níger'),(660,'ca',566,'NG','Nigèria'),(661,'ca',570,'NU','Niue'),(662,'ca',574,'NF','Illa Norfolk'),(663,'ca',580,'MP','Illes Marianes del Nord'),(664,'ca',578,'NO','Noruega'),(665,'ca',512,'OM','Oman'),(666,'ca',586,'PK','Pakistan'),(667,'ca',585,'PW','Palau'),(668,'ca',275,'PS','Territori Palestí, Ocupat'),(669,'ca',591,'PA','Panamà'),(670,'ca',598,'PG','Papua Nova Guinea'),(671,'ca',600,'PY','Paraguai'),(672,'ca',604,'PE','Perú'),(673,'ca',608,'PH','Filipines'),(674,'ca',612,'PN','Pitcairn'),(675,'ca',616,'PL','Polònia'),(676,'ca',620,'PT','Portugal'),(677,'ca',630,'PR','Puerto Rico'),(678,'ca',634,'QA','Qatar'),(679,'ca',638,'RE','Reunion'),(680,'ca',642,'RO','Romania'),(681,'ca',643,'RU','Federació Russa'),(682,'ca',646,'RW','Rwanda'),(683,'ca',652,'BL','Saint Barthélemy'),(684,'ca',654,'SH','Saint Helena, Ascension i Tristan da Cunha'),(685,'ca',659,'KN','Saint Christopher i Nevis'),(686,'ca',662,'LC','Saint Lucia'),(687,'ca',663,'MF','Saint Martin (part francesa)'),(688,'ca',666,'PM','Saint Pierre i Miquelon'),(689,'ca',670,'VC','Saint Vincent i les Grenadines'),(690,'ca',882,'WS','Samoa'),(691,'ca',674,'SM','San Marino'),(692,'ca',678,'ST','Sao Tome i Príncipe'),(693,'ca',682,'SA','Aràbia Saudita'),(694,'ca',686,'SN','Senegal'),(695,'ca',688,'RS','Sèrbia'),(696,'ca',690,'SC','Seychelles'),(697,'ca',694,'SL','Sierra Leone'),(698,'ca',702,'SG','Singapur'),(699,'ca',534,'SX','Sint Maarten (part neerlandesa)'),(700,'ca',703,'SK','Eslovàquia'),(701,'ca',705,'SI','Eslovènia'),(702,'ca',90,'SB','Illes Salomó'),(703,'ca',706,'SO','Somàlia'),(704,'ca',710,'ZA','Sudàfrica'),(705,'ca',239,'GS','Illes Geòrgia del Sud i Sandwich del Sud'),(706,'ca',724,'ES','Espanya'),(707,'ca',144,'LK','Sri Lanka'),(708,'ca',729,'SD','Sudan'),(709,'ca',740,'SR','Surinam'),(710,'ca',728,'SS','South Sudan'),(711,'ca',744,'SJ','Svalbard i Jan Mayen'),(712,'ca',748,'SZ','Swazilàndia'),(713,'ca',752,'SE','Suècia'),(714,'ca',756,'CH','Suïssa'),(715,'ca',760,'SY','República Àrab Síria'),(716,'ca',158,'TW','Taiwan, Província de Xina'),(717,'ca',762,'TJ','Tadjikistan'),(718,'ca',834,'TZ','Tanzània, República Unida de'),(719,'ca',764,'TH','Tailàndia'),(720,'ca',626,'TL','Timor Oriental'),(721,'ca',768,'TG','Togo'),(722,'ca',772,'TK','Tokelau'),(723,'ca',776,'TO','Tonga'),(724,'ca',780,'TT','Trinitat i Tobago'),(725,'ca',788,'TN','Tunísia'),(726,'ca',792,'TR','Turquia'),(727,'ca',795,'TM','Turkmenistan'),(728,'ca',796,'TC','Illes Turks i Caicos'),(729,'ca',798,'TV','Tuvalu'),(730,'ca',800,'UG','Uganda'),(731,'ca',804,'UA','Ucraïna'),(732,'ca',784,'AE','Unió dels Emirats Àrabs'),(733,'ca',826,'GB','Regne Unit'),(734,'ca',840,'US','Estats Units'),(735,'ca',581,'UM','Illes Perifèriques Menors dels EUA'),(736,'ca',858,'UY','Uruguai'),(737,'ca',860,'UZ','Uzbekistan'),(738,'ca',548,'VU','Vanuatu'),(739,'ca',862,'VE','Veneçuela, República Bolívariana de'),(740,'ca',704,'VN','Vietnam'),(741,'ca',92,'VG','Illes Verges, Britàniques'),(742,'ca',850,'VI','Illes Verges, EUA'),(743,'ca',876,'WF','Wallis i Futuna'),(744,'ca',732,'EH','Sàhara Occidental'),(745,'ca',887,'YE','Iemen'),(746,'ca',894,'ZM','Zàmbia'),(747,'ca',716,'ZW','Zimbabwe');
/*!40000 ALTER TABLE `countries` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `customers`
--

DROP TABLE IF EXISTS `customers`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `customers` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `TIN` varchar(16) NOT NULL DEFAULT '',
  `Name` varchar(255) NOT NULL DEFAULT '',
  `Address` varchar(255) DEFAULT '',
  `City` varchar(255) DEFAULT '',
  `Postalcode` varchar(80) DEFAULT '',
  `Country` varchar(80) DEFAULT '',
  `WEB` varchar(255) DEFAULT '',
  `Phone` varchar(80) DEFAULT '',
  `Fax` varchar(80) DEFAULT '',
  `MobilePhone` varchar(80) DEFAULT '',
  `eMail` varchar(80) DEFAULT '',
  `Text` varchar(255) DEFAULT '',
  `Discount` decimal(4,2) DEFAULT '0.00',
  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
  `TicketOrder` tinyint(1) DEFAULT '0',
  `Points` int(11) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Code` (`TIN`),
  UNIQUE KEY `TIN` (`TIN`),
  KEY `Name` (`Name`(19))
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `customers`
--

LOCK TABLES `customers` WRITE;
/*!40000 ALTER TABLE `customers` DISABLE KEYS */;
INSERT INTO `customers` VALUES (1,'1','CLIENTE UNO','CIF VETE A SABER QUE.','','','','','','','','','',0.00,0,0,0),(2,'2','CLIENTE DOS','NIF 545454545454','','','','','','','','','',0.00,0,0,0);
/*!40000 ALTER TABLE `customers` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `datalogger`
--

DROP TABLE IF EXISTS `datalogger`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `datalogger` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` int(11) NOT NULL DEFAULT '0',
  `Scale` int(11) NOT NULL DEFAULT '0',
  `DatTim` datetime NOT NULL DEFAULT '1970-01-01 00:00:00',
  `Message` text NOT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=31 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `datalogger`
--

LOCK TABLES `datalogger` WRITE;
/*!40000 ALTER TABLE `datalogger` DISABLE KEYS */;
INSERT INTO `datalogger` VALUES (1,133,1,'2018-10-10 14:54:36','Network database updated at version 111'),(2,133,1,'2018-10-10 14:54:37','Network database updated at version 111'),(3,131,1,'2018-10-10 14:54:42','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(4,131,1,'2018-10-10 15:03:33','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(5,131,1,'2018-10-10 15:23:38','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(6,131,1,'2018-10-10 15:34:21','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(7,131,1,'2018-10-10 15:49:48','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(8,131,1,'2018-10-10 13:55:19','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(9,131,1,'2018-10-10 14:01:55','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(10,131,1,'2018-10-11 08:01:11','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(11,131,1,'2018-10-11 11:32:09','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(12,131,1,'2018-12-12 09:37:41','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(13,131,1,'2018-12-12 09:39:11','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(14,131,1,'2019-02-19 18:07:55','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(15,131,1,'2019-02-21 15:48:52','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(16,131,1,'2019-02-21 17:52:02','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(17,131,1,'2019-02-21 18:45:48','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(18,131,1,'2019-02-25 10:40:51','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(19,131,1,'2019-02-25 17:59:19','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(20,131,1,'2019-02-26 10:39:40','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(21,131,1,'2019-02-27 09:52:37','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(22,131,1,'2019-02-27 18:07:08','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(23,131,1,'2019-02-27 18:13:08','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(24,131,1,'2019-02-27 18:17:20','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(25,131,1,'2019-02-27 18:18:33','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(26,131,1,'2019-02-27 18:28:20','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(27,131,1,'2019-02-28 11:16:16','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(28,131,1,'2019-02-28 17:06:33','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(29,131,1,'2019-02-28 19:15:59','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false'),(30,131,1,'2019-02-28 19:34:03','ClearOlderDays call with Type=\'T\', Days:366, ResetNumber:false ClearDaily:false');
/*!40000 ALTER TABLE `datalogger` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `ean128`
--

DROP TABLE IF EXISTS `ean128`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `ean128` (
  `Counter` int(11) NOT NULL DEFAULT '0',
  `Code1` varchar(32) DEFAULT '',
  `Code2` varchar(32) DEFAULT '',
  `Code3` varchar(32) DEFAULT '',
  `Code4` varchar(32) DEFAULT '',
  `Code5` varchar(32) DEFAULT '',
  `Code6` varchar(32) DEFAULT '',
  `Code7` varchar(32) DEFAULT '',
  `Code8` varchar(32) DEFAULT '',
  `Code9` varchar(32) DEFAULT '',
  `Code10` varchar(32) DEFAULT '',
  PRIMARY KEY (`Counter`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `ean128`
--

LOCK TABLES `ean128` WRITE;
/*!40000 ALTER TABLE `ean128` DISABLE KEYS */;
/*!40000 ALTER TABLE `ean128` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `erasetickets`
--

DROP TABLE IF EXISTS `erasetickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `erasetickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` char(1) DEFAULT 'T',
  `DatTim` datetime DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Type` (`Type`)
) ENGINE=InnoDB AUTO_INCREMENT=24 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `erasetickets`
--

LOCK TABLES `erasetickets` WRITE;
/*!40000 ALTER TABLE `erasetickets` DISABLE KEYS */;
INSERT INTO `erasetickets` VALUES (21,'T','2019-02-28 19:14:00'),(22,'P','2016-10-26 10:29:17'),(23,'S','2015-11-16 08:23:51');
/*!40000 ALTER TABLE `erasetickets` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertEraseTickets` AFTER INSERT ON `EraseTickets`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `EraseTickets`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = latin1 */ ;
/*!50003 SET character_set_results = latin1 */ ;
/*!50003 SET collation_connection  = latin1_swedish_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `TicketsErased` AFTER UPDATE ON `EraseTickets`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `EraseTickets`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `families`
--

DROP TABLE IF EXISTS `families`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `families` (
  `Code` int(11) NOT NULL,
  `Name` varchar(54) NOT NULL DEFAULT '',
  `Icon` varchar(127) DEFAULT '',
  `DefaultIcon` varchar(127) DEFAULT '',
  `IsBovine` tinyint(1) DEFAULT '0',
  `Label` varchar(64) NOT NULL,
  PRIMARY KEY (`Code`),
  KEY `Name` (`Name`(19))
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `families`
--

LOCK TABLES `families` WRITE;
/*!40000 ALTER TABLE `families` DISABLE KEYS */;
INSERT INTO `families` VALUES (1,'DPT. 1','','',0,''),(201,'VACUNO','','',0,''),(202,'PORCINO','','',0,''),(203,'OVINO Y CAPRINO','','',0,''),(204,'EQUINO','','',0,''),(205,'AVES Y CAZA','','',0,''),(206,'ESPECIALIDADES NAVIDAD','','',0,''),(207,'HUEVOS','','',0,''),(208,'ELABORADOS','','',0,''),(209,'CODIGOS GENERALES','','',0,''),(210,'EMBUTIDOS','','',0,''),(301,'CHARCUTERIA FIAMBRES','','',0,''),(302,'QUESOS','','',0,''),(310,'LOTES','','',0,''),(501,'FRUTAS','3413.png','',0,''),(502,'VERDURAS Y HORTALIZAS','3914.png','',0,''),(503,'CODIGOS GENERALES','','',0,''),(505,'FRUTOS SECOS','43900.png','',0,''),(601,'PAN','','',0,''),(602,'BOLLERÍA','','',0,''),(603,'PASTELERÍA Y REPOSTERÍA','','',0,''),(604,'CODIGOS GENERALES','','',0,''),(701,'PESCADO FRESCO','','',0,''),(702,'MARISCO','','',0,''),(703,'CODIGOS GENERALES','','',0,''),(2001,'BASE ARROZ','','',0,''),(2002,'PANADERÖA','','',0,''),(2003,'BASE VERDURAS','','',0,''),(2004,'BASE CARNE','','',0,''),(2005,'BASE PESCADO','','',0,''),(2006,'PIZZAS','','',0,''),(2008,'ENSALADILLA','','',0,''),(2009,'MASA','','',0,''),(5303,'CHARCUTERÍA','','',0,''),(5305,'PESCADERÍA','','',0,''),(5306,'PANADERÍA','','',0,'');
/*!40000 ALTER TABLE `families` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertFamily` AFTER INSERT ON `Families`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckFamilies`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateFamily` AFTER UPDATE ON `Families`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckFamilies`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteFamily` AFTER DELETE ON `Families`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckFamilies`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `hangedmarks`
--

DROP TABLE IF EXISTS `hangedmarks`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `hangedmarks` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Identifier` char(1) NOT NULL DEFAULT '0',
  `Description` varchar(127) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Identifier` (`Identifier`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `hangedmarks`
--

LOCK TABLES `hangedmarks` WRITE;
/*!40000 ALTER TABLE `hangedmarks` DISABLE KEYS */;
/*!40000 ALTER TABLE `hangedmarks` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `hcustomerorders`
--

DROP TABLE IF EXISTS `hcustomerorders`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `hcustomerorders` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Number` varchar(12) NOT NULL,
  `Section` decimal(3,0) DEFAULT '0',
  `UseSection` tinyint(1) DEFAULT '0',
  `OrderDate` date NOT NULL DEFAULT '0000-00-00',
  `Customer` varchar(127) NOT NULL DEFAULT '',
  `Phone` varchar(20) NOT NULL DEFAULT '',
  `DeliveryDate` date NOT NULL DEFAULT '0000-00-00',
  `DeliveryHour` time NOT NULL DEFAULT '00:00:00',
  `PreparationDateLimit` date NOT NULL DEFAULT '0000-00-00',
  `PreparationHourLimit` time NOT NULL DEFAULT '00:00:00',
  `StartDateTime` datetime NOT NULL DEFAULT '1970-01-01 00:00:00',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) NOT NULL DEFAULT '0',
  `AddedLines` int(11) DEFAULT '0',
  `NetAmount` decimal(10,2) DEFAULT '0.00',
  `TotalWeight` decimal(12,3) DEFAULT '0.000',
  `TotalUnits` int(11) DEFAULT '0',
  `Packages` int(11) NOT NULL DEFAULT '1',
  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `Status` char(1) NOT NULL DEFAULT '0',
  `Captured` char(1) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `hctor_idx` (`Number`,`Section`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `hcustomerorders`
--

LOCK TABLES `hcustomerorders` WRITE;
/*!40000 ALTER TABLE `hcustomerorders` DISABLE KEYS */;
/*!40000 ALTER TABLE `hcustomerorders` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `HCtordInserted` AFTER INSERT ON `HCustomerOrders`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckCustomerOrders` = TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `HCtordUpdated` AFTER UPDATE ON `HCustomerOrders`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckCustomerOrders` = TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `HCtordDeleted` AFTER DELETE ON `HCustomerOrders`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckCustomerOrders` = TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `htickets`
--

DROP TABLE IF EXISTS `htickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `htickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `IdNetHticket` int(11) DEFAULT '0',
  `NetStat` char(1) DEFAULT '',
  `Captured` char(1) DEFAULT '',
  `CapturedDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `tckType` char(1) DEFAULT NULL,
  `Status` char(1) DEFAULT NULL,
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `TurnPrefix` char(1) DEFAULT '0',
  `Turn` int(10) unsigned DEFAULT '0',
  `TurnId` decimal(20,0) DEFAULT '0',
  `TurnDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `TurnOnTime` tinyint(1) NOT NULL DEFAULT '0',
  `SectionItem` int(11) DEFAULT '0',
  `UseSectionItem` tinyint(1) DEFAULT '0',
  `NumLines` int(11) DEFAULT '0',
  `AddedLines` int(11) DEFAULT '0',
  `GrossAmount` decimal(10,2) DEFAULT '0.00',
  `NetAmount` decimal(10,2) DEFAULT '0.00',
  `TotalWeight` decimal(12,3) DEFAULT '0.000',
  `TotalUnits` int(11) DEFAULT '0',
  `NegativeLines` int(11) DEFAULT '0',
  `NegativeAmount` decimal(10,2) DEFAULT '0.00',
  `CanceledLines` int(11) DEFAULT '0',
  `CanceledAmount` decimal(10,2) DEFAULT '0.00',
  `EAN13` char(13) NOT NULL DEFAULT '',
  `Code128` varchar(32) NOT NULL DEFAULT '',
  `StartDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `TimesPrinted` tinyint(3) unsigned NOT NULL DEFAULT '0',
  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadFrom` int(11) DEFAULT '0',
  `ReloadScale` int(11) NOT NULL DEFAULT '0',
  `ReloadDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadTicketNumber` int(11) NOT NULL DEFAULT '0',
  `ReloadtckType` char(1) DEFAULT '',
  `ReloadTo` int(11) DEFAULT '0',
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `Points` int(11) NOT NULL DEFAULT '0',
  `PointsExtra` int(11) NOT NULL DEFAULT '0',
  `PointsDiscount` int(11) NOT NULL DEFAULT '0',
  `PointsAmount` decimal(10,2) NOT NULL DEFAULT '0.00',
  `PointsChanged` int(11) NOT NULL DEFAULT '0',
  `OrderNumber` varchar(12) DEFAULT '',
  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
  `Invoice` varchar(20) NOT NULL DEFAULT '',
  `InvoiceDate` date DEFAULT '0000-00-00',
  `tckDiscount` decimal(4,2) DEFAULT '0.00',
  `ParkedNumber` int(11) DEFAULT '0',
  `NeedCopy` tinyint(1) DEFAULT '0',
  `CashStatus` char(1) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT '-',
  `CashOperId` int(11) DEFAULT '0',
  `CashAmount` decimal(10,2) DEFAULT '0.00',
  `CashIP` int(10) unsigned NOT NULL DEFAULT '0',
  `Checksum` varchar(1024) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  KEY `Vendor` (`Vendor`),
  KEY `Invoice` (`Invoice`),
  KEY `datim_idx` (`DatTim`),
  KEY `OrderNumber` (`OrderNumber`)
) ENGINE=InnoDB AUTO_INCREMENT=20 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `htickets`
--

LOCK TABLES `htickets` WRITE;
/*!40000 ALTER TABLE `htickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `htickets` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `HTicketsInsertB` BEFORE INSERT ON `HTickets`
FOR EACH ROW BEGIN
  IF(New.`OrderNumber`!=0) THEN
    SET @vnd = (SELECT `Vendor` FROM `HCustomerOrders` WHERE (`HCustomerOrders`.`Number`=New.`OrderNumber` AND `HCustomerOrders`.`Section`=New.`SectionItem`));
    IF(@vnd=0) THEN 
      UPDATE `HCustomerOrders` SET `HCustomerOrders`.`Vendor`=NEW.`Vendor`, `HCustomerOrders`.`DatTim`=NOW(), `HCustomerOrders`.`Scale`=NEW.`Scale`,
        `HCustomerOrders`.`StartDateTime`=IF(`HCustomerOrders`.`Status`=1, NOW(), `HCustomerOrders`.`StartDateTime`), `HCustomerOrders`.`Status`=2
        WHERE (`HCustomerOrders`.`Number`=NEW.`OrderNumber` AND `HCustomerOrders`.`Section`=New.`SectionItem`) AND ((`HCustomerOrders`.`Status`>0) AND (`HCustomerOrders`.`Status`<=2));
    ELSE
      UPDATE `PCScaleErrorTag` SET x=1;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `HTicketsInsert` AFTER INSERT ON `HTickets`
FOR EACH ROW BEGIN
  IF (NEW.`NetStat`='N') AND (NEW.`Counter`= GLOBAL_GET("current_counter")) THEN
    REPLACE INTO `LocalVendorChanges` VALUES ((NEW.`Vendor`*10)+1, NEW.`Vendor`,'1');
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `HTicketsUpdateB` BEFORE UPDATE ON `HTickets`
FOR EACH ROW BEGIN
  IF(NEW.`NetStat`='N') THEN
    IF ((OLD.`GrossAmount`!=NEW.`GrossAmount`) OR (OLD.`Customer`!=NEW.`Customer`)) THEN
      SET NEW.`PointsDiscount`=0;
      SET NEW.`PointsAmount`= 0;
    END IF;
    IF (OLD.`tckDiscount`!=NEW.`tckDiscount`) THEN
      SET NEW.`PointsDiscount`=0;
      SET NEW.`PointsAmount`= 0;
      SET NEW.`NetAmount`=CAST(NEW.`GrossAmount`*((100-NEW.`tckDiscount`)/100) AS DECIMAL(10,2));
    ELSE
      IF (OLD.`PointsDiscount`!=NEW.`PointsDiscount`) THEN
        SET @lypoints = (SELECT IFNULL(`ParamValue`,0) FROM `CfgSystem` WHERE `ParamKey`='loyalty-points-discount-currency');
        SET NEW.`PointsAmount`= IF(@lypoints,CAST(NEW.`PointsDiscount`/@lypoints AS DECIMAL(10,2)),0);
        SET NEW.`NetAmount`=CAST(NEW.`GrossAmount`*((100-NEW.`tckDiscount`)/100) AS DECIMAL(10,2))-NEW.`PointsAmount`;
      END IF;
    END IF;
    IF (OLD.`NetAmount`!=NEW.`NetAmount`) THEN
      SET @points = (SELECT IFNULL(`ParamValue`,0) FROM `CfgSystem` WHERE `ParamKey`='loyalty-currency-give-point');
      SET NEW.`Points`=IF(@points,FLOOR(NEW.`NetAmount`/@points),0);
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `HTicketsUpdate` AFTER UPDATE ON `HTickets`
FOR EACH ROW BEGIN

  IF(NEW.`ReloadTo`!=0) AND (OLD.`Status`!= NEW.`Status`) THEN
    UPDATE `LTickets` SET `LTickets`.`ReloadTo`=NEW.`ReloadTo` WHERE `LTickets`.`IdHTicket`=OLD.`Id`;
  END IF;

  IF((NEW.`NetStat`='N') AND (NEW.`Counter`= GLOBAL_GET("current_counter"))) THEN
    REPLACE INTO `LocalVendorChanges` VALUES ((NEW.`Vendor`*10)+1, NEW.`Vendor`,'1');

    UPDATE `LocalHTickets` SET `Status`=NEW.`Status`, `TimesPrinted`=NEW.`TimesPrinted`,
      `Captured`=IF(NEW.`Captured`!='',NEW.`Captured`,`Captured`), `CapturedDateTime`=IF(NEW.`Captured`!='',NEW.`CapturedDateTime`,`CapturedDateTime`)
      WHERE `IdNetHticket`=NEW.`Id`;
  END IF;
  IF(OLD.`Number`!= NEW.`Number`) THEN
    UPDATE `LTickets` SET `LTickets`.`PrintScale`=NEW.`Scale`, `LTickets`.`PrintDateTime`=NEW.`DatTim`, `LTickets`.`PrintTicketNumber`=NEW.`Number`,
      `LTickets`.`PrinttckType`=OLD.`tckType`, `LTickets`.`TicketDiscount`=OLD.`tckDiscount` WHERE `LTickets`.`IdHTicket`=OLD.`Id`;
  END IF;

  IF(OLD.`Status`!= NEW.`Status`) THEN
    IF(NEW.`Status`='C') THEN

      INSERT INTO `TotalDaily` (`Date`, `Type`, `NetStat`, `Tickets`, `TotalLines`, `TotalAmount`, `TotalDiscount`, `TotalWeight`, `TotalUnits`,
        `NegativeLines`, `NegativeAmount`, `CanceledLines`, `CanceledAmount`, `Invoices`, `InvoiceAmount`)
        VALUES (DATE(NEW.`DatTim`), NEW.`tckType`, NEW.`NetStat`, 1, NEW.`NumLines`, NEW.`NetAmount`, (NEW.`GrossAmount`-NEW.`NetAmount`), NEW.`TotalWeight`, NEW.`TotalUnits`,
        NEW.`NegativeLines`, NEW.`NegativeAmount`, NEW.`CanceledLines`, NEW.`CanceledAmount`, IF(NEW.`Invoice`!='',1,0), IF(NEW.`Invoice`!='',NEW.`NetAmount`,0) )
        ON DUPLICATE KEY UPDATE `TotalDaily`.`PrintChgCounter`=`TotalDaily`.`PrintChgCounter`+`TotalDaily`.`Printed`, `TotalDaily`.`Printed`=0,
          `TotalDaily`.`Tickets`=`TotalDaily`.`Tickets`+1,`TotalDaily`.`TotalLines`=`TotalDaily`.`TotalLines`+NEW.`NumLines`,`TotalAmount`=`TotalAmount`+NEW.`NetAmount`,
          `TotalDiscount`=`TotalDiscount`+VALUES(`TotalDiscount`), `TotalWeight`=`TotalWeight`+NEW.`TotalWeight`, `TotalUnits`=`TotalUnits`+NEW.`TotalUnits`,
          `NegativeLines`=`NegativeLines`+NEW.`NegativeLines`, `NegativeAmount`=`NegativeAmount`+NEW.`NegativeAmount`,
          `CanceledLines`=`CanceledLines`+NEW.`CanceledLines`, `CanceledAmount`=`CanceledAmount`+NEW.`CanceledAmount`,
          `Invoices`=`Invoices`+VALUES(`Invoices`), `InvoiceAmount`=`InvoiceAmount`+VALUES(`InvoiceAmount`);

      INSERT INTO `TotalPerpetual` (`Date`, `Type`, `NetStat`, `Tickets`, `TotalLines`, `TotalAmount`, `TotalDiscount`, `TotalWeight`, `TotalUnits`,
        `NegativeLines`, `NegativeAmount`, `CanceledLines`, `CanceledAmount`, `Invoices`, `InvoiceAmount`)
        VALUES (DATE(NEW.`DatTim`), NEW.`tckType`, NEW.`NetStat`, 1, NEW.`NumLines`, NEW.`NetAmount`, (NEW.`GrossAmount`-NEW.`NetAmount`), NEW.`TotalWeight`, NEW.`TotalUnits`,
        NEW.`NegativeLines`, NEW.`NegativeAmount`, NEW.`CanceledLines`, NEW.`CanceledAmount`, IF(NEW.`Invoice`!='',1,0), IF(NEW.`Invoice`!='',NEW.`NetAmount`,0) )
        ON DUPLICATE KEY UPDATE 
          `TotalPerpetual`.`Tickets`=`TotalPerpetual`.`Tickets`+1,`TotalPerpetual`.`TotalLines`=`TotalPerpetual`.`TotalLines`+NEW.`NumLines`,`TotalAmount`=`TotalAmount`+NEW.`NetAmount`,
          `TotalDiscount`=`TotalDiscount`+VALUES(`TotalDiscount`), `TotalWeight`=`TotalWeight`+NEW.`TotalWeight`, `TotalUnits`=`TotalUnits`+NEW.`TotalUnits`,
          `NegativeLines`=`NegativeLines`+NEW.`NegativeLines`, `NegativeAmount`=`NegativeAmount`+NEW.`NegativeAmount`,
          `CanceledLines`=`CanceledLines`+NEW.`CanceledLines`, `CanceledAmount`=`CanceledAmount`+NEW.`CanceledAmount`,
          `Invoices`=`Invoices`+VALUES(`Invoices`), `InvoiceAmount`=`InvoiceAmount`+VALUES(`InvoiceAmount`);

      IF(NEW.`GrossAmount`!=0) THEN
        INSERT INTO `TotalVAT` (`Date`, `Type`, `NetStat`, `Code`, `Percentage`, `Amount`, `Base`, `Tax`)
          SELECT DATE(NEW.`DatTim`), NEW.`tckType`, NEW.`NetStat`, `V_VAT`.`VATCode`, `V_VAT`.`VATPercent`, `V_VAT`.`Amount`,
          CAST((`V_VAT`.`Amount`*100)/(100+`VATPercent`) AS DECIMAL(10,2)) AS `Base`, `V_VAT`.`Amount` - CAST((`V_VAT`.`Amount`*100)/(100+`VATPercent`) AS DECIMAL(10,2)) AS Tax
          FROM (SELECT `VATCode`, `VATPercent`, CAST(SUM(`LTickets`.`Amount`)*(NEW.`NetAmount`/NEW.`GrossAmount`) AS DECIMAL(10,2)) AS `Amount`
            FROM `LTickets` WHERE `LTickets`.`IdHTicket`=NEW.`Id` AND `LTickets`.`lineType`=1 GROUP BY `LTickets`.`VATCode`) AS `V_VAT`
            ON DUPLICATE KEY UPDATE `TotalVAT`.`Amount`=`TotalVAT`.`Amount`+VALUES(`Amount`), `TotalVAT`.`Base`=`TotalVAT`.`Base`+VALUES(`Base`),
            `TotalVAT`.`Tax`=`TotalVAT`.`Tax`+VALUES(`Tax`);
      END IF;

      UPDATE `Stock`,
        (SELECT `Item`, SUM(IF(`Saleform`=1,IF(`Price`<0,-`Weight`,`Weight`),0)) AS `lwght`, SUM(IF(`Saleform`=0,IF(`Price`<0,-`Weight`,`Weight`),0)) AS `lunits` FROM `LTickets` WHERE (`IdHTicket`=NEW.`Id` AND ((`LTickets`.`lineType`&1)=1)) GROUP BY `Item`) AS `VT_L`
        SET `Weight`=IF((`Weight`-`VT_L`.`lwght`)>0, (`Weight`-`VT_L`.`lwght`),`Weight`), `Stock`.`Units`=IF((`Stock`.`Units`-`VT_L`.`lunits`)>0, (`Stock`.`Units`-`VT_L`.`lunits`),`Stock`.`Units`) WHERE `Code` = `VT_L`.`Item`;

      INSERT INTO `TotalPayTickets` (`Date`, `Type`, `NetStat`, `PayMethod`, `Delivered`)
        (SELECT DATE(NEW.`DatTim`), NEW.`tckType`, NEW.`NetStat`,`PayMethod`, `Delivered` FROM `PayTickets` WHERE `PayTickets`.`IdHTicket`=NEW.`Id`)
        ON DUPLICATE KEY UPDATE `TotalPayTickets`.`Delivered`=`TotalPayTickets`.`Delivered`+VALUES(`Delivered`);

      UPDATE `Customers` SET `Customers`.`Points`=GREATEST(`Customers`.`Points`+(NEW.`Points`+NEW.`PointsExtra`)-(NEW.`PointsDiscount`+NEW.`PointsChanged`),0) WHERE `Customers`.`TIN`=NEW.`Customer`;

      IF(NEW.`tckType`!='U') THEN


        IF(OLD.`SimplifiedInvoice`!=0) THEN
          INSERT INTO `Invoices` (`Number`, `Date`, `Customer`, `Type`) VALUES (OLD.`Invoice`, OLD.`InvoiceDate`, OLD.`Customer`, '0');
          INSERT INTO `InvoiceDetails` (`IdInvoice`, `IdHTicket`) VALUES (LAST_INSERT_ID(), OLD.`Id`);
        END IF;

        INSERT INTO `TotalScaleFamilies` (`Date`, `Scale`, `Family`, `TotalWeight`, `TotalUnits`)
          SELECT CONCAT(YEAR(NEW.`DatTim`),'-',MONTH(NEW.`DatTim`), '-01'), NEW.`Scale`,
          `Family`, SUM(IF(`Saleform`=1,`Weight`,0)), SUM(IF(`Saleform`=0,`Weight`,0)) FROM `LTickets` WHERE (`IdHTicket`=NEW.`Id` AND `LTickets`.`lineType`=1) GROUP BY `Family`
          ON DUPLICATE KEY UPDATE `TotalScaleFamilies`.`PrintChgCounter`=`TotalScaleFamilies`.`PrintChgCounter`+`TotalScaleFamilies`.`Printed`, `TotalScaleFamilies`.`Printed`=0,
          `TotalWeight`=`TotalWeight`+ VALUES(`TotalWeight`), `TotalUnits`=`TotalUnits`+ VALUES(`TotalUnits`);

      END IF;

    END IF;


    IF(OLD.`OrderNumber`!='') THEN

      IF((NEW.`Status`='P') OR (NEW.`Status`='C') OR (NEW.`Status`='X')) THEN
        UPDATE `HCustomerOrders` SET `HCustomerOrders`.`Vendor`=0, `HCustomerOrders`.`DatTim`=NOW()
          WHERE (`HCustomerOrders`.`Number`=OLD.`OrderNumber` AND `HCustomerOrders`.`Section`=OLD.`SectionItem`) AND `HCustomerOrders`.`Status`=2 AND `HCustomerOrders`.`Vendor`=NEW.`Vendor`;
      END IF;
    END IF;

    IF(OLD.`Status`='C') THEN


      UPDATE `TotalDaily` SET `PrintChgCounter`=`PrintChgCounter`+`Printed`, `Printed`=0,
        `Tickets`=`Tickets`-1, `TotalLines`=`TotalLines`-OLD.`NumLines`,`TotalAmount`=`TotalAmount`-OLD.`NetAmount`,
        `TotalDiscount`=`TotalDiscount`-(OLD.`GrossAmount`-OLD.`NetAmount`), `TotalWeight`=`TotalWeight`-OLD.`TotalWeight`, `TotalUnits`=`TotalUnits`-OLD.`TotalUnits`,
        `NegativeLines`=`NegativeLines`-OLD.`NegativeLines`, `NegativeAmount`=`NegativeAmount`-OLD.`NegativeAmount`,
        `CanceledLines`=`CanceledLines`-OLD.`CanceledLines`, `CanceledAmount`=`CanceledAmount`-OLD.`CanceledAmount`,
        `Invoices`=`Invoices`-IF(OLD.`Invoice`!='',1,0), `InvoiceAmount`=`InvoiceAmount`-IF(OLD.`Invoice`!='',OLD.`NetAmount`,0)
        WHERE `Date`=DATE(OLD.`DatTim`) AND `Type`=OLD.`tckType` AND `NetStat`=OLD.`NetStat`;

      UPDATE `TotalPerpetual` SET 
        `Tickets`=`Tickets`-1, `TotalLines`=`TotalLines`-OLD.`NumLines`,`TotalAmount`=`TotalAmount`-OLD.`NetAmount`,
        `TotalDiscount`=`TotalDiscount`-(OLD.`GrossAmount`-OLD.`NetAmount`), `TotalWeight`=`TotalWeight`-OLD.`TotalWeight`, `TotalUnits`=`TotalUnits`-OLD.`TotalUnits`,
        `NegativeLines`=`NegativeLines`-OLD.`NegativeLines`, `NegativeAmount`=`NegativeAmount`-OLD.`NegativeAmount`,
        `CanceledLines`=`CanceledLines`-OLD.`CanceledLines`, `CanceledAmount`=`CanceledAmount`-OLD.`CanceledAmount`,
        `Invoices`=`Invoices`-IF(OLD.`Invoice`!='',1,0), `InvoiceAmount`=`InvoiceAmount`-IF(OLD.`Invoice`!='',OLD.`NetAmount`,0)
        WHERE `Type`=OLD.`tckType` AND `NetStat`=OLD.`NetStat`;


      IF(NEW.`GrossAmount`!=0) THEN
        UPDATE `TotalVAT`,
          (SELECT `VATCode`, `VATPercent`, CAST(SUM(`LTickets`.`Amount`)*(OLD.`NetAmount`/OLD.`GrossAmount`) AS DECIMAL(10,2)) AS `Amount`
          FROM `LTickets` WHERE `LTickets`.`IdHTicket`=OLD.`Id` AND `LTickets`.`lineType`=1 GROUP BY `LTickets`.`VATCode`) AS `V_VAT`
          SET `TotalVAT`.`Amount`=`TotalVAT`.`Amount`-`V_VAT`.`Amount`,
          `TotalVAT`.`Base`=`TotalVAT`.`Base`-(CAST((`V_VAT`.`Amount`*100)/(100+`VATPercent`) AS DECIMAL(10,2))),
          `TotalVAT`.`Tax`=`TotalVAT`.`Tax`-(`V_VAT`.`Amount` - CAST((`V_VAT`.`Amount`*100)/(100+`VATPercent`) AS DECIMAL(10,2)))
          WHERE `Date`=DATE(OLD.`DatTim`) AND `Type`=OLD.`tckType` AND `NetStat`=OLD.`NetStat` AND `Code`=`V_VAT`.`VATCode`;
      END IF;

      UPDATE `Stock`,
        (SELECT `Item`, SUM(IF(`Saleform`=1,IF(`Price`<0,-`Weight`,`Weight`),0)) AS `lwght`, SUM(IF(`Saleform`=0,IF(`Price`<0,-`Weight`,`Weight`),0)) AS `lunits` FROM `LTickets` WHERE (`IdHTicket`=NEW.`Id`  AND ((`LTickets`.`lineType`&1)=1)) GROUP BY `Item`) AS `VT_L`
        SET `Weight`=IF((`Weight`+`VT_L`.`lwght`)>0, (`Weight`+`VT_L`.`lwght`),`Weight`), `Stock`.`Units`=IF((`Stock`.`Units`+`VT_L`.`lunits`)>0, (`Stock`.`Units`+`VT_L`.`lunits`),`Stock`.`Units`) WHERE `Code` = `VT_L`.`Item`;

        UPDATE `TotalPayTickets`,
          (SELECT `PayMethod`, SUM(`Delivered`) AS `Delivered` FROM `PayTickets` WHERE `PayTickets`.`IdHTicket`=OLD.`Id` GROUP BY `PayMethod`) AS `V_PT`
          SET `TotalPayTickets`.`Delivered`=`TotalPayTickets`.`Delivered`-`V_PT`.`Delivered`
          WHERE `Date`=DATE(OLD.`DatTim`) AND `Type`=OLD.`tckType` AND `NetStat`=OLD.`NetStat` AND `TotalPayTickets`.`PayMethod`=`V_PT`.`PayMethod`;

        UPDATE `Customers` SET `Customers`.`Points`=GREATEST(`Customers`.`Points`-(OLD.`Points`+OLD.`PointsExtra`)+(OLD.`PointsDiscount`+OLD.`PointsChanged`),0) WHERE `Customers`.`TIN`=NEW.`Customer`;

      IF (OLD.`tckType`!='U') THEN


        UPDATE `TotalScaleFamilies`,
          (SELECT CONCAT(YEAR(NEW.`DatTim`),'-',MONTH(NEW.`DatTim`), '-01') AS `fdt`,
          `Family`, SUM(IF(`Saleform`=1,`Weight`,0)) AS `fwght`, SUM(IF(`Saleform`=0,`Weight`,0)) AS `funits` FROM `LTickets` WHERE (`IdHTicket`=NEW.`Id` AND `LTickets`.`lineType`=1) GROUP BY `Family`) AS `VT_F`
          SET `TotalWeight`=`TotalWeight`-`VT_F`.`fwght` , `TotalUnits`=`TotalUnits`+ `VT_F`.`funits` WHERE (`Date`=`VT_F`.`fdt` AND `Scale`=NEW.`Scale` AND `TotalScaleFamilies`.`Family`=`VT_F`.`Family`);

      IF(OLD.`OrderNumber`!='') THEN
        UPDATE `HCustomerOrders` SET `HCustomerOrders`.`AddedLines`=`HCustomerOrders`.`AddedLines`-OLD.`AddedLines`, `HCustomerOrders`.`NetAmount`=`HCustomerOrders`.`NetAmount`-OLD.`NetAmount`,
          `HCustomerOrders`.`TotalWeight`=`HCustomerOrders`.`TotalWeight`-OLD.`TotalWeight`, `HCustomerOrders`.`TotalUnits`=`HCustomerOrders`.`TotalUnits`-OLD.`TotalUnits`, `HCustomerOrders`.`DatTim`=NOW()
          WHERE (`HCustomerOrders`.`Number`=OLD.`OrderNumber` AND `HCustomerOrders`.`Section`=OLD.`SectionItem`) AND `HCustomerOrders`.`Status`=2;
      END IF;

      END IF;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `HTicketsDelete` AFTER DELETE ON `HTickets`
FOR EACH ROW BEGIN
  IF (OLD.`Counter`= GLOBAL_GET("current_counter")) THEN
    REPLACE INTO `LocalVendorChanges` VALUES ((OLD.`Vendor`*10)+1, OLD.`Vendor`,'1');
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `inventory`
--

DROP TABLE IF EXISTS `inventory`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `inventory` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Number` int(11) DEFAULT '0',
  `PrintDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `Code` int(11) DEFAULT NULL,
  `SectionItem` int(11) DEFAULT '0',
  `Name` varchar(127) NOT NULL,
  `DateTime` datetime DEFAULT NULL,
  `Weight` decimal(5,3) DEFAULT NULL,
  `Units` int(11) DEFAULT NULL,
  `Status` smallint(6) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `inventory`
--

LOCK TABLES `inventory` WRITE;
/*!40000 ALTER TABLE `inventory` DISABLE KEYS */;
/*!40000 ALTER TABLE `inventory` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `invoicedetails`
--

DROP TABLE IF EXISTS `invoicedetails`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `invoicedetails` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `IdInvoice` int(11) DEFAULT NULL,
  `IdHTicket` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `IdInvoice` (`IdInvoice`,`IdHTicket`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `invoicedetails`
--

LOCK TABLES `invoicedetails` WRITE;
/*!40000 ALTER TABLE `invoicedetails` DISABLE KEYS */;
/*!40000 ALTER TABLE `invoicedetails` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `invoices`
--

DROP TABLE IF EXISTS `invoices`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `invoices` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Number` varchar(20) NOT NULL DEFAULT '',
  `OriginalNumber` varchar(20) NOT NULL DEFAULT '',
  `Date` date DEFAULT '0000-00-00',
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `Type` char(1) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `invoices`
--

LOCK TABLES `invoices` WRITE;
/*!40000 ALTER TABLE `invoices` DISABLE KEYS */;
/*!40000 ALTER TABLE `invoices` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `RemoveInvoice` AFTER DELETE ON `Invoices`
FOR EACH ROW BEGIN
  IF (@delete_tickets_from_invoices=1) THEN
    DELETE `InvoiceDetails`, `HTickets` FROM `InvoiceDetails` INNER JOIN `HTickets` WHERE IdHTicket=HTickets.Id AND IdInvoice=OLD.`Id`;
  ELSE
    UPDATE `HTickets` LEFT JOIN `InvoiceDetails` ON HTickets.Id=InvoiceDetails.IdHTicket SET HTickets.Invoice=DEFAULT(HTickets.Invoice), HTickets.InvoiceDate=DEFAULT(HTickets.InvoiceDate) WHERE IdInvoice=OLD.`Id`;
    DELETE FROM `InvoiceDetails` WHERE IdInvoice=OLD.`Id`;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `itemgroupkeys`
--

DROP TABLE IF EXISTS `itemgroupkeys`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemgroupkeys` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `NumKey` decimal(2,0) NOT NULL,
  `AboveSection` decimal(2,0) DEFAULT '0',
  `Section` decimal(2,0) DEFAULT '0',
  `SubSection` decimal(2,0) DEFAULT '0',
  `Family` decimal(2,0) DEFAULT '0',
  `SubFamily` decimal(2,0) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `idx_` (`NumKey`,`AboveSection`,`Section`,`SubSection`,`Family`,`SubFamily`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemgroupkeys`
--

LOCK TABLES `itemgroupkeys` WRITE;
/*!40000 ALTER TABLE `itemgroupkeys` DISABLE KEYS */;
/*!40000 ALTER TABLE `itemgroupkeys` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertItemGroupKeys` AFTER INSERT ON `ItemGroupKeys`
FOR EACH ROW BEGIN
  INSERT INTO `LocalItemGroupKeysChanges` (`NumKey`) VALUES (NEW.`NumKey`) ON DUPLICATE KEY UPDATE `NumKey`=VALUES(`NumKey`);
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateItemGroupKeys` AFTER UPDATE ON `ItemGroupKeys`
FOR EACH ROW BEGIN
  INSERT INTO `LocalItemGroupKeysChanges` (`NumKey`) VALUES (OLD.`NumKey`) ON DUPLICATE KEY UPDATE `NumKey`=VALUES(`NumKey`);
  INSERT INTO `LocalItemGroupKeysChanges` (`NumKey`) VALUES (NEW.`NumKey`) ON DUPLICATE KEY UPDATE `NumKey`=VALUES(`NumKey`);
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteItemGroupKeys` AFTER DELETE ON `ItemGroupKeys`
FOR EACH ROW BEGIN
  INSERT INTO `LocalItemGroupKeysChanges` (`NumKey`) VALUES (OLD.`NumKey`) ON DUPLICATE KEY UPDATE `NumKey`=VALUES(`NumKey`);
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `itemgroupkeystext`
--

DROP TABLE IF EXISTS `itemgroupkeystext`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemgroupkeystext` (
  `NumKey` decimal(2,0) NOT NULL,
  `Text` varchar(64) DEFAULT '',
  `Icon` varchar(127) DEFAULT '',
  PRIMARY KEY (`NumKey`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemgroupkeystext`
--

LOCK TABLES `itemgroupkeystext` WRITE;
/*!40000 ALTER TABLE `itemgroupkeystext` DISABLE KEYS */;
/*!40000 ALTER TABLE `itemgroupkeystext` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertItemGroupKeysText` AFTER INSERT ON `ItemGroupKeysText`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckItemGroupKeysText`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateItemGroupKeysText` AFTER UPDATE ON `ItemGroupKeysText`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckItemGroupKeysText`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteItemGroupKeysText` AFTER DELETE ON `ItemGroupKeysText`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckItemGroupKeysText`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `itemhabitual`
--

DROP TABLE IF EXISTS `itemhabitual`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemhabitual` (
  `Position` tinyint(3) unsigned NOT NULL DEFAULT '1',
  `Code` int(11) NOT NULL,
  PRIMARY KEY (`Position`),
  KEY `fk_1` (`Code`),
  CONSTRAINT `itemhabitual_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemhabitual`
--

LOCK TABLES `itemhabitual` WRITE;
/*!40000 ALTER TABLE `itemhabitual` DISABLE KEYS */;
/*!40000 ALTER TABLE `itemhabitual` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemHabitualInserted` AFTER INSERT ON `ItemHabitual`
FOR EACH ROW BEGIN
  SET @toffer := IFNULL (GLOBAL_GET("table-offer"),0);
  SET @thabitual := IFNULL (GLOBAL_GET("table-habitual"),0);
  IF (@thabitual!=0) THEN
    IF ((SELECT IF(`Items`.`PLU`!=0 AND (`Items`.`Offer`=0 OR @toffer=0) ,0,1) FROM `Items` WHERE `Items`.`Code`=NEW.`Code`)=0) THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (GLOBAL_GET("current_group"),@thabitual,NEW.`Position`,NEW.`Code`)
        ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemHabitualUpdated` AFTER UPDATE ON `ItemHabitual`
FOR EACH ROW BEGIN
  SET @toffer := IFNULL (GLOBAL_GET("table-offer"),0);
  SET @thabitual := IFNULL (GLOBAL_GET("table-habitual"),0);
  IF (@thabitual!=0) THEN
    INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (GLOBAL_GET("current_group"),@thabitual,OLD.`Position`,0)
      ON DUPLICATE KEY UPDATE `Code`=0;
    IF ((SELECT IF(`Items`.`PLU`!=0 AND (`Items`.`Offer`=0 OR @toffer=0) ,0,1) FROM `Items` WHERE `Items`.`Code`=NEW.`Code`)=0) THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (GLOBAL_GET("current_group"),@thabitual,NEW.`Position`,NEW.`Code`)
        ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemHabitualDeleted` AFTER DELETE ON `ItemHabitual`
FOR EACH ROW BEGIN
  SET @thabitual := IFNULL (GLOBAL_GET("table-habitual"),0);
  IF (@thabitual!=0) THEN
    INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (GLOBAL_GET("current_group"),@thabitual,OLD.`Position`,0)
    ON DUPLICATE KEY UPDATE `Code`=0;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `itemlinkcommon`
--

DROP TABLE IF EXISTS `itemlinkcommon`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemlinkcommon` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` int(11) NOT NULL,
  `NumCard` int(11) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `commonlink_idx` (`Code`,`NumCard`),
  CONSTRAINT `ilc_ibfk1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemlinkcommon`
--

LOCK TABLES `itemlinkcommon` WRITE;
/*!40000 ALTER TABLE `itemlinkcommon` DISABLE KEYS */;
/*!40000 ALTER TABLE `itemlinkcommon` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `itempricerates`
--

DROP TABLE IF EXISTS `itempricerates`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itempricerates` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` int(11) NOT NULL,
  `Rate` decimal(1,0) NOT NULL DEFAULT '1',
  `Price` decimal(8,2) NOT NULL DEFAULT '0.00',
  `Step` decimal(5,3) NOT NULL DEFAULT '0.000',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `idx_rate` (`Code`,`Rate`,`Step`),
  KEY `Rate` (`Rate`),
  CONSTRAINT `itempricerates_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE,
  CONSTRAINT `itempricerates_ibfk_2` FOREIGN KEY (`Rate`) REFERENCES `pricerates` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itempricerates`
--

LOCK TABLES `itempricerates` WRITE;
/*!40000 ALTER TABLE `itempricerates` DISABLE KEYS */;
/*!40000 ALTER TABLE `itempricerates` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `items`
--

DROP TABLE IF EXISTS `items`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `items` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` int(11) NOT NULL,
  `Name` varchar(127) NOT NULL DEFAULT '',
  `ShortName` varchar(64) DEFAULT '',
  `Text` text,
  `Text1` text,
  `Text2` text,
  `Text3` text,
  `Text4` text,
  `Text5` text,
  `Family` int(11) DEFAULT '0',
  `SubFamily` decimal(2,0) DEFAULT '0',
  `FamilyOrder` int(11) DEFAULT '0',
  `OnKey` tinyint(1) DEFAULT '1',
  `AboveSection` decimal(2,0) DEFAULT '0',
  `Section` decimal(3,0) DEFAULT '0',
  `SubSection` decimal(2,0) DEFAULT '0',
  `WorkingMode` varchar(20) DEFAULT '',
  `Price` decimal(8,2) DEFAULT '0.00',
  `Offer` tinyint(1) DEFAULT '0',
  `Discount` decimal(4,2) DEFAULT '0.00',
  `SaleForm` smallint(6) DEFAULT '1',
  `PLU` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `VAT` decimal(1,0) unsigned DEFAULT '0',
  `EAN13` varchar(13) DEFAULT '',
  `Expiring` smallint(6) DEFAULT '0',
  `BestBefore` smallint(6) DEFAULT '0',
  `Tare` decimal(7,3) DEFAULT '0.000',
  `Label` varchar(64) DEFAULT '',
  `Label2` varchar(64) DEFAULT '-',
  `Nutrition` int(11) DEFAULT '0',
  `EAN13Print` varchar(12) DEFAULT NULL,
  `PacketExt` char(1) NOT NULL DEFAULT '-',
  `IsFrozen` tinyint(1) DEFAULT '0',
  `Icon` varchar(127) DEFAULT '',
  `Icon2` varchar(127) DEFAULT '',
  `QLevel` varchar(16) DEFAULT '',
  `Mixed` smallint(6) DEFAULT '0',
  `MixedWeight` decimal(7,3) DEFAULT '0.000',
  `CommonCard` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Code` (`Code`),
  KEY `Name` (`Name`(19)),
  KEY `ide_fmor` (`Family`,`FamilyOrder`)
) ENGINE=InnoDB AUTO_INCREMENT=12485 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `items`
--

LOCK TABLES `items` WRITE;
/*!40000 ALTER TABLE `items` DISABLE KEYS */;
INSERT INTO `items` VALUES (10696,0,'','',NULL,NULL,NULL,NULL,NULL,NULL,0,0,999,0,0,0,0,'',0.00,0,0.00,1,0,0,0,'',0,0,0.000,'','-',0,NULL,'-',0,'','','',0,0.000,0),(11704,1127,'COPA DE UVA NAVIDAD','COPA DE UVA NAVIDAD','','Origen:alicante;Categoria:extra;Variedad:aledo','','','','',501,0,1,0,0,0,0,'',1.00,0,0.00,1,217,4,0,'',0,0,0.000,'','-',0,'','-',0,'1127.png','','',0,0.000,0),(11705,1176,'HUESO RAPE FRESCO','HUESO RAPE FRESCO','','','','','','',701,0,3,1,0,0,0,'',7.60,0,0.00,1,53,5,0,'',0,0,0.000,'','-',0,'','-',0,'1176.png','','',0,0.000,0),(11706,1207,'CABUT','CABUT','','','','','','',701,0,2,1,0,0,0,'',5.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'1207.png','','',0,0.000,0),(11707,1616,'PANCETA IBERICA SALADA','PANCETA IBERICA SALADA','','','','','','',301,0,14,1,0,0,0,'',4.99,0,0.00,1,340,3,0,'',0,0,0.000,'','-',0,'','-',0,'1616.png','','',0,0.000,0),(11708,1617,'HUESOS DE JAMON','HUESOS DE JAMON','','','','','','',301,0,2,1,0,0,0,'',1.99,0,0.00,1,543,3,0,'',30,0,0.000,'','-',0,'','-',0,'1617.png','','',0,0.000,0),(11709,1621,'SALMON MEDIO O ENTERO 3-5','SALMON MEDIO O ENTERO 3-5','','','','','','',701,0,3,1,0,0,0,'',10.95,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'1621.png','','',0,0.000,0),(11710,1677,'SOLOMILLO DE CERDO BANDEJ','SOLOMILLO DE CERDO BANDEJ','','','','','','',202,0,12,1,0,0,0,'',8.99,0,0.00,1,133,2,0,'',0,0,0.000,'','-',0,'','-',0,'1677.png','','',0,0.000,0),(11711,2046,'CHULETAS CORDERO PAL-RIÑO','CHULETAS CORDERO PAL-RIÑO','','','','','','',203,0,10,1,0,0,0,'',16.98,0,0.00,1,560,2,0,'',0,0,0.000,'','-',0,'','-',0,'2046.png','','',0,0.000,0),(11712,2081,'CHULETAS CORDERO PIERNA B','CHULETAS CORDERO PIERNA B','','','','','','',203,0,9,1,0,0,0,'',14.98,0,0.00,1,559,2,0,'',0,0,0.000,'','-',0,'','-',0,'2081.png','','',0,0.000,0),(11713,2276,'CHOPPED DE PAVO S/GLUTEN','CHOPPED DE PAVO S/GLUTEN','','','','','','',301,0,6,1,0,0,0,'',7.90,0,0.00,1,52,3,0,'',30,0,0.000,'','-',0,'','-',0,'2276.png','','',0,0.000,0),(11714,2279,'PECHUGA PAVO','PECHUGA PAVO','','','','','','',301,0,5,1,0,0,0,'',9.99,0,0.00,1,36,3,0,'',15,0,0.000,'','-',0,'','-',0,'2279.png','','',0,0.000,0),(11715,2355,'ATUN','ATUN','','','','','','',701,0,4,1,0,0,0,'',12.90,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'2355.png','','',0,0.000,0),(11716,2467,'CALAMAR EXTRA','CALAMAR EXTRA','','','','','','',702,0,1,1,0,0,0,'',12.90,0,0.00,1,62,5,0,'',0,0,0.000,'','-',0,'','-',0,'2467.png','','',0,0.000,0),(11717,2487,'CABEZA LOMO CON HUESO','CABEZA LOMO CON HUESO','','','','','','',202,0,2,1,0,0,0,'',5.50,0,0.00,1,37,2,0,'',0,0,0.000,'','-',0,'','-',0,'2487.png','','',0,0.000,0),(11718,2514,'CHORIZO TRADICION ROJO EX','CHORIZO TRADICION ROJO EX','','','','','','',301,0,4,1,0,0,0,'',11.95,0,0.00,1,23,3,0,'',0,0,0.000,'','-',0,'','-',0,'2514.png','','',0,0.000,0),(11719,2515,'PULPO GORDO FRESCO 1KG. A','PULPO GORDO FRESCO 1KG. A','','','','','','',702,0,2,1,0,0,0,'',15.50,0,0.00,1,65,5,0,'',0,0,0.000,'','-',0,'','-',0,'2515.png','','',0,0.000,0),(11720,2525,'COSTILLAS DE CERDO','COSTILLAS DE CERDO','','','','','','',202,0,6,1,0,0,0,'',6.80,0,0.00,1,48,2,0,'',0,0,0.000,'','-',0,'','-',0,'2525.png','','',0,0.000,0),(11721,2537,'MORTADELA ACEITUNAS','MORTADELA ACEITUNAS','','','','','','',301,0,3,1,0,0,0,'',5.45,0,0.00,1,17,3,0,'',30,0,0.000,'','-',0,'','-',0,'2537.png','','',0,0.000,0),(11722,2550,'JAMON COCIDO C/PATA C/HUE','JAMON COCIDO C/PATA C/HUE','','','','','','',301,0,13,1,0,0,0,'',11.95,0,0.00,1,321,3,0,'',15,0,0.000,'','-',0,'','-',0,'2550.png','','',0,0.000,0),(11723,2607,'TOCINO IBERICO FRESCO','TOCINO IBERICO FRESCO','','','','','','',202,0,17,1,0,0,0,'',3.99,0,0.00,1,547,2,0,'',0,0,0.000,'','-',0,'','-',0,'2607.png','','',0,0.000,0),(11724,2655,'BERENJENA NEGRA','BERENJENA NEGRA','','Origen:España;Categoria:primera;Calibre:g (300-400 g.);Variedad:negra','','','','',502,0,3,1,0,0,0,'',2.79,0,0.00,1,13,4,0,'',0,0,0.000,'','-',0,'','-',0,'2655.png','','',0,0.000,0),(11725,2696,'BABILLA VACUNO','BABILLA VACUNO','','','','','','',201,0,10,1,0,0,0,'',12.98,0,0.00,1,302,2,0,'',0,0,0.000,'','-',0,'','-',0,'2696.png','','',0,0.000,0),(11726,2714,'CIGALAS GORDAS GRANEL 0','CIGALAS GORDAS GRANEL 0','','','','','','',702,0,3,1,0,0,0,'',29.90,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'2714.png','','',0,0.000,0),(11727,2716,'PIMIENTO ITALIANO','PIMIENTO ITALIANO','','Origen:España;Categoria:1º;Calibre:g;Variedad:largo poniente','','','','',502,0,7,1,0,0,0,'',1.99,0,0.00,1,26,4,0,'',0,0,0.000,'','-',0,'','-',0,'2716.png','','',0,0.000,0),(11728,2722,'POMELOS','POMELOS','','Origen:España;Categoria:i;Calibre:88/102;Variedad:star ruby. trat. imazalil','','','','',501,0,26,1,0,0,0,'',1.49,0,0.00,1,82,4,0,'',0,0,0.000,'','-',0,'','-',0,'2722.png','','',0,0.000,0),(11729,2745,'LIMONES','LIMONES','','Origen:España;Categoria:i;Calibre:58/67;Variedad:primafliori;Tratamiento:tratado con imazalil','','','','',501,0,2,1,0,0,0,'',1.69,0,0.00,1,3,4,0,'',0,0,0.000,'','-',0,'','-',0,'2745.png','','',0,0.000,0),(11730,2938,'1/2 COL LOMBARDA','1/2 COL LOMBARDA','','Origen:España;Categoria:1ª;Variedad:roja','','','','',502,0,3,0,0,0,0,'',1.00,0,0.00,1,105,4,0,'',0,0,0.000,'','-',0,'','-',0,'2938.png','','',0,0.000,0),(11731,2987,'1/2 COL','1/2 COL','','Origen:España','','','','',502,0,4,0,0,0,0,'',1.00,0,0.00,1,104,4,0,'',0,0,0.000,'','-',0,'','-',0,'2987.png','','',0,0.000,0),(11732,3155,'CAQUIS','CAQUIS','','Origen:España;Categoria:i;Variedad:rojo brillante persimon;Calibre:20/22','','','','',501,0,15,1,0,0,0,'',1.99,0,0.00,1,55,4,0,'',0,0,0.000,'','-',0,'','-',0,'3155.png','','',0,0.000,0),(11733,3195,'OSTRAS BANDEJA','OSTRAS BANDEJA','','','','','','',702,0,4,1,0,0,0,'',5.70,0,0.00,1,92,5,0,'',0,0,0.000,'','-',0,'','-',0,'3195.png','','',0,0.000,0),(11734,3276,'BONIATO ROJO','BONIATO ROJO','','Origen:España;Categoria:primera;Calibre:+300;Variedad:rojo','','','','',502,0,21,1,0,0,0,'',1.85,0,0.00,1,85,4,0,'',0,0,0.000,'','-',0,'','-',0,'3276.png','','',0,0.000,0),(11735,3277,'CHULETAS LOMO CERDO','CHULETAS LOMO CERDO','','','','','','',202,0,5,1,0,0,0,'',5.99,0,0.00,1,36,2,0,'',0,0,0.000,'','-',0,'','-',0,'3277.png','','',0,0.000,0),(11736,3406,'SOLOMILLO CERDO','SOLOMILLO CERDO','','','','','','',202,0,4,1,0,0,0,'',8.99,0,0.00,1,39,2,0,'',0,0,0.000,'','-',0,'','-',0,'3406.png','','',0,0.000,0),(11737,3481,'CHIRIMOYAS','CHIRIMOYAS','','Origen:España;Categoria:i;Calibre:25;Variedad:fina de jete','','','','',501,0,22,1,0,0,0,'',3.49,0,0.00,1,72,4,0,'',0,0,0.000,'','-',0,'','-',0,'3481.png','','',0,0.000,0),(11738,3628,'GAMBA ALISTADA PEQ.','GAMBA ALISTADA PEQ.','','','','','','',702,0,3,1,0,0,0,'',10.90,0,0.00,1,70,5,0,'',0,0,0.000,'','-',0,'','-',0,'3628.png','','',0,0.000,0),(11739,3630,'LOMO EMBUCHADO','LOMO EMBUCHADO','','Ingredientes:lomo de cerdo,sal,especias,LACTOSA,azucar;Conservadores:E-252,E-250;Antioxidante:E-301;Condiciones de conservacion:sitio fresco y seco','','','','',301,0,7,1,0,0,0,'',15.75,0,0.00,1,60,3,0,'',0,0,0.000,'','-',0,'','-',0,'3630.png','','',0,0.000,0),(11740,3637,'BERENJENA RAYADA','BERENJENA RAYADA','','Origen:España;Categoria:i;Calibre:280+;Variedad:rayada','','','','',502,0,4,1,0,0,0,'',3.25,0,0.00,1,14,4,0,'',0,0,0.000,'','-',0,'','-',0,'3637.png','','',0,0.000,0),(11741,3689,'QUESO AZUL HOJAS','QUESO AZUL HOJAS','','','','','','',302,0,9,1,0,0,0,'',11.95,0,0.00,1,158,3,0,'',0,0,0.000,'','-',0,'','-',0,'3689.png','','',0,0.000,0),(11742,3763,'PERA BLANQUILLA','PERA BLANQUILLA','','Origen:España;Categoria:i;Calibre:55+;Variedad:blanquilla','','','','',501,0,20,1,0,0,0,'',1.59,0,0.00,1,65,4,0,'',0,0,0.000,'','-',0,'','-',0,'3763.png','','',0,0.000,0),(11743,3776,'PLATANOS','PLATANOS','','Origen:canarias (España);Categoria:i;Variedad:cavendish;Calibre:+14','','','','',501,0,6,1,0,0,0,'',2.29,0,0.00,1,9,4,0,'',0,0,0.000,'','-',0,'','-',0,'3776.png','','',0,0.000,0),(11744,3803,'MANDARINAS EXTRA','MANDARINAS EXTRA','','Origen:España;Categoria:i;Calibre:2/3;Variedad:oronules','','','','',501,0,19,1,0,0,0,'',1.59,0,0.00,1,62,4,0,'',0,0,0.000,'','-',0,'','-',0,'3803.png','','',0,0.000,0),(11745,3806,'QUESO CORTE','QUESO CORTE','','','','','','',302,0,7,1,0,0,0,'',15.95,0,0.00,1,129,3,0,'',0,0,0.000,'','-',0,'','-',0,'3806.png','','',0,0.000,0),(11746,3830,'QUESO OVEJA CURADO','QUESO OVEJA CURADO','','','','','','',302,0,4,1,0,0,0,'',13.75,0,0.00,1,77,3,0,'',0,0,0.000,'','-',0,'','-',0,'3830.png','','',0,0.000,0),(11747,3884,'UVA NEGRA','UVA NEGRA','','Origen:peru;Categoria:i;Variedad:red globe;Calibre:l','','','','',501,0,24,1,0,0,0,'',3.59,0,0.00,1,78,4,0,'',0,0,0.000,'','-',0,'','-',0,'3884.png','','',0,0.000,0),(11748,3888,'QUESO FINAS HIERBAS','QUESO FINAS HIERBAS','','','','','','',302,0,6,1,0,0,0,'',16.90,0,0.00,1,121,3,0,'',30,0,0.000,'','-',0,'','-',0,'3888.png','','',0,0.000,0),(11749,3889,'QUESO NUECES','QUESO NUECES','','','','','','',302,0,8,1,0,0,0,'',16.90,0,0.00,1,140,3,0,'',21,0,0.000,'','-',0,'','-',0,'3889.png','','',0,0.000,0),(11750,3907,'COSTILLA VACUNO','COSTILLA VACUNO','','','','','','',201,0,6,1,0,0,0,'',6.95,0,0.00,1,98,2,0,'',0,0,0.000,'','-',0,'','-',0,'3907.png','','',0,0.000,0),(11751,3914,'PIMIENTO VERDE','PIMIENTO VERDE','','Origen:España;Categoria:i;Calibre:g;Variedad:lamuyo verde','','','','',502,0,19,1,0,0,0,'',2.29,0,0.00,1,79,4,0,'',0,0,0.000,'','-',0,'','-',0,'3914.png','','',0,0.000,0),(11752,3922,'APIO VERDE','APIO VERDE','','','','','','',502,0,8,0,0,0,0,'',1.10,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'3922.png','','',0,0.000,0),(11753,3929,'PUERROS BANDEJA','PUERROS BANDEJA','','Origen:España;Categoria:primera','','','','',502,0,9,0,0,0,0,'',2.25,0,0.00,1,120,4,0,'',0,0,0.000,'','-',0,'','-',0,'3929.png','','',0,0.000,0),(11754,3932,'PERA ALEJANDRINA','PERA ALEJANDRINA','','Origen:España;Categoria:i;Calibre:70/75;Variedad:alejandrina','','','','',501,0,29,1,0,0,0,'',1.89,0,0.00,1,95,4,0,'',0,0,0.000,'','-',0,'','-',0,'3932.png','','',0,0.000,0),(11755,3938,'NARANJA','NARANJA','','Origen:España;Categoria:i;Calibre:73/84;Variedad:navelina','','','','',501,0,5,1,0,0,0,'',1.19,0,0.00,1,8,4,0,'',0,0,0.000,'','-',0,'','-',0,'3938.png','','',0,0.000,0),(11756,3940,'AGUACATES','AGUACATES','','Origen:Perú;Categoria:i;Calibre:14-16;Variedad:hass','','','','',501,0,1,1,0,0,0,'',4.99,0,0.00,1,1,4,0,'',0,0,0.000,'','-',0,'','-',0,'3940.png','','',0,0.000,0),(11757,3946,'CALABACIN','CALABACIN','','Origen:España;Categoria:primera;Calibre:g (14);Variedad:verde','','','','',502,0,5,1,0,0,0,'',2.49,0,0.00,1,15,4,0,'',0,0,0.000,'','-',0,'','-',0,'3946.png','','',0,0.000,0),(11758,3948,'GRANADAS','GRANADAS','','Origen:España;Categoria:i;Variedad:mollar;Calibre:70/90','','','','',501,0,16,1,0,0,0,'',1.49,0,0.00,1,56,4,0,'',0,0,0.000,'','-',0,'','-',0,'3948.png','','',0,0.000,0),(11759,3965,'PEPINO ESPAÑOL CORTO','PEPINO ESPAÑOL CORTO','','Origen:España;Categoria:primera;Calibre:l;Variedad:corto','','','','',502,0,18,1,0,0,0,'',1.75,0,0.00,1,73,4,0,'',0,0,0.000,'','-',0,'','-',0,'3965.png','','',0,0.000,0),(11760,3976,'LOMO CINTA','LOMO CINTA','','','','','','',202,0,5,1,0,0,0,'',7.99,0,0.00,1,40,2,0,'',0,0,0.000,'','-',0,'','-',0,'3976.png','','',0,0.000,0),(11761,3999,'RABOS CERDO','RABOS CERDO','','','','','','',202,0,10,1,0,0,0,'',3.99,0,0.00,1,55,2,0,'',0,0,0.000,'','-',0,'','-',0,'3999.png','','',0,0.000,0),(11762,4033,'PIMIENTO ROJO','PIMIENTO ROJO','','Origen:España;Categoria:1;Calibre:g;Variedad:lamuyo','','','','',502,0,8,1,0,0,0,'',2.29,0,0.00,1,27,4,0,'',0,0,0.000,'','-',0,'','-',0,'4033.png','','',0,0.000,0),(11763,4034,'CALABAZA','CALABAZA','','Origen:España;Categoria:i;Variedad:pasty delice;Calibre:1 kg','','','','',502,0,22,1,0,0,0,'',1.15,0,0.00,1,87,4,0,'',0,0,0.000,'','-',0,'','-',0,'4034.png','','',0,0.000,0),(11764,4036,'CHULETAS PIERNA CORDERO','CHULETAS PIERNA CORDERO','','','','','','',203,0,3,1,0,0,0,'',13.95,0,0.00,1,21,2,0,'',0,0,0.000,'','-',0,'','-',0,'4036.png','','',0,0.000,0),(11765,4065,'TOMATE CANARIO','TOMATE CANARIO','','Origen:España;Categoria:i;Calibre:m;Variedad:long life','','','','',502,0,9,1,0,0,0,'',1.59,0,0.00,1,29,4,0,'',0,0,0.000,'','-',0,'','-',0,'4065.png','','',0,0.000,0),(11766,4089,'ALCACHOFAS','ALCACHOFAS','','Origen:España;Categoria:i;Calibre:5/6;Variedad:tudela','','','','',502,0,20,1,0,0,0,'',2.25,0,0.00,1,83,4,0,'',0,0,0.000,'','-',0,'','-',0,'4089.png','','',0,0.000,0),(11767,4134,'(GRANEL) TARTA MANZANA','(GRANEL) TARTA MANZANA','','','','','','',603,0,1,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',60,0,0.000,'','-',0,'','-',0,'4134.png','','',0,0.000,0),(11768,4142,'(GRANEL) BRACITOS AZUCAR','(GRANEL) BRACITOS AZUCAR','','','','','','',603,0,2,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',60,0,0.000,'','-',0,'','-',0,'4142.png','','',0,0.000,0),(11769,4143,'(GRANEL) BRACITOS TRUFA C','(GRANEL) BRACITOS TRUFA C','','','','','','',603,0,3,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',60,0,0.000,'','-',0,'','-',0,'4143.png','','',0,0.000,0),(11770,4153,'(GRANEL) CAÑAS CREMA CHOC','(GRANEL) CAÑAS CREMA CHOC','','','','','','',603,0,4,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',60,0,0.000,'','-',0,'','-',0,'4153.png','','',0,0.000,0),(11771,4179,'(GRANEL) PALMERA CHOC.GRA','(GRANEL) PALMERA CHOC.GRA','','','','','','',603,0,5,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',60,0,0.000,'','-',0,'','-',0,'4179.png','','',0,0.000,0),(11772,4218,'PECHO Y FALDA CORDERO','PECHO Y FALDA CORDERO','','','','','','',203,0,4,1,0,0,0,'',4.98,0,0.00,1,27,2,0,'',0,0,0.000,'','-',0,'','-',0,'4218.png','','',0,0.000,0),(11773,4220,'CEBOLLA MALLA 2 KG','CEBOLLA MALLA 2 KG','','','','','','',502,0,16,0,0,0,0,'',1.55,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'4220.png','','',0,0.000,0),(11774,4268,'CUELLO CORDERO','CUELLO CORDERO','','','','','','',203,0,5,1,0,0,0,'',8.20,0,0.00,1,26,2,0,'',0,0,0.000,'','-',0,'','-',0,'4268.png','','',0,0.000,0),(11775,4328,'PANCETA DE CERDO','PANCETA DE CERDO','','','','','','',202,0,9,1,0,0,0,'',6.50,0,0.00,1,49,2,0,'',0,0,0.000,'','-',0,'','-',0,'4328.png','','',0,0.000,0),(11776,4412,'MAGRO GUISAR','MAGRO GUISAR','','','','','','',202,0,11,1,0,0,0,'',5.99,0,0.00,1,82,2,0,'',0,0,0.000,'','-',0,'','-',0,'4412.png','','',0,0.000,0),(11777,4436,'PAVO NACIONAL 4,5-5,5 KG','PAVO NACIONAL 4,5-5,5 KG','','','','','','',205,0,7,1,0,0,0,'',4.40,0,0.00,1,17,2,0,'',0,0,0.000,'','-',0,'','-',0,'4436.png','','',0,0.000,0),(11778,4449,'ENTRECOT DE VACUNO','ENTRECOT DE VACUNO','','','','','','',201,0,2,1,0,0,0,'',18.25,0,0.00,1,77,2,0,'',0,0,0.000,'','-',0,'','-',0,'4449.png','','',0,0.000,0),(11779,4488,'JAMONCITOS PAVO','JAMONCITOS PAVO','','','','','','',205,0,2,1,0,0,0,'',3.75,0,0.00,1,19,2,0,'',0,0,0.000,'','-',0,'','-',0,'4488.png','','',0,0.000,0),(11780,4492,'CARNE VACUNO DE GUISAR','CARNE VACUNO DE GUISAR','','','','','','',201,0,5,1,0,0,0,'',7.50,0,0.00,1,97,2,0,'',0,0,0.000,'','-',0,'','-',0,'4492.png','','',0,0.000,0),(11781,4524,'GARRETA-DOBLADA DE VACUNO','GARRETA-DOBLADA DE VACUNO','','','','','','',201,0,5,1,0,0,0,'',8.98,0,0.00,1,92,2,0,'',0,0,0.000,'','-',0,'','-',0,'4524.png','','',0,0.000,0),(11782,4701,'LENGUAS DE VACUNO','LENGUAS DE VACUNO','','','','','','',201,0,8,1,0,0,0,'',7.65,0,0.00,1,106,2,0,'',0,0,0.000,'','-',0,'','-',0,'4701.png','','',0,0.000,0),(11783,4713,'NABETA BANDEJA','NABETA BANDEJA','','','','','','',502,0,17,0,0,0,0,'',2.69,0,0.00,1,215,4,0,'',0,0,0.000,'','-',0,'','-',0,'4713.png','','',0,0.000,0),(11784,4714,'CHIRIVIA BANDEJA','CHIRIVIA BANDEJA','','','','','','',502,0,18,0,0,0,0,'',2.69,0,0.00,1,216,4,0,'',0,0,0.000,'','-',0,'','-',0,'4714.png','','',0,0.000,0),(11785,4807,'PEPINO HOLANDES','PEPINO HOLANDES','','Origen:España;Categoria:primera;Calibre:l','','','','',502,0,15,1,0,0,0,'',2.20,0,0.00,1,67,4,0,'',0,0,0.000,'','-',0,'','-',0,'4807.png','','',0,0.000,0),(11786,4808,'KIWI','KIWI','','Origen:Nueva Zelanda;Categoria:primera;Calibre:33/36;Variedad:hayward green','','','','',501,0,18,1,0,0,0,'',3.99,0,0.00,1,59,4,0,'',0,0,0.000,'','-',0,'','-',0,'4808.png','','',0,0.000,0),(11787,4809,'PIÑA','PIÑA','','Origen:Costa Rica;Categoria:extra;Variedad:cavendish;Calibre:6/7','','','','',501,0,27,1,0,0,0,'',1.19,0,0.00,1,88,4,0,'',0,0,0.000,'','-',0,'','-',0,'4809.png','','',0,0.000,0),(11788,4810,'NAPICOL BANDEJA','NAPICOL BANDEJA','','','','','','',502,0,20,0,0,0,0,'',2.69,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'4810.png','','',0,0.000,0),(11789,4820,'MANZANA ROYAL GALA','MANZANA ROYAL GALA','','Origen:España;Categoria:i;Calibre:70/+;Variedad:royal gala','','','','',501,0,14,1,0,0,0,'',1.59,0,0.00,1,48,4,0,'',0,0,0.000,'','-',0,'','-',0,'4820.png','','',0,0.000,0),(11790,4870,'PATATA BOLSA 3KG','PATATA BOLSA 3KG','','','','','','',502,0,21,0,0,0,0,'',3.70,0,0.00,1,218,4,0,'',0,0,0.000,'','-',0,'','-',0,'4870.png','','',0,0.000,0),(11791,4874,'RODABALLO','RODABALLO','','','','','','',701,0,1,1,0,0,0,'',14.90,0,0.00,1,19,5,0,'',0,0,0.000,'','-',0,'','-',0,'4874.png','','',0,0.000,0),(11792,4905,'GAMBON 20-30 ALTA MAR','GAMBON 20-30 ALTA MAR','','','','','','',702,0,10,1,0,0,0,'',10.90,0,0.00,1,247,5,0,'',0,0,0.000,'','-',0,'','-',0,'4905.png','','',0,0.000,0),(11793,4957,'SALMON MEDIO O ENTERO 1-3','SALMON MEDIO O ENTERO 1-3','','','','','','',701,0,6,1,0,0,0,'',8.50,0,0.00,1,45,5,0,'',0,0,0.000,'','-',0,'','-',0,'4957.png','','',0,0.000,0),(11794,5061,'CALAMAR PATAGONICO','CALAMAR PATAGONICO','','','','','','',702,0,6,1,0,0,0,'',6.95,0,0.00,1,121,5,0,'',0,0,0.000,'','-',0,'','-',0,'5061.png','','',0,0.000,0),(11795,5092,'EMPERADOR TACOS','EMPERADOR TACOS','','','','','','',701,0,4,1,0,0,0,'',8.50,0,0.00,1,54,5,0,'',0,0,0.000,'','-',0,'','-',0,'5092.png','','',0,0.000,0),(11796,5147,'BURGER DE POLLO Y CERDO','BURGER DE POLLO Y CERDO','','','','','','',208,0,1,1,0,0,0,'',6.98,0,0.00,1,277,2,0,'',3,0,0.000,'','-',0,'','-',0,'5147.png','','',0,0.000,0),(11797,5171,'GARRETA CORDERO','GARRETA CORDERO','','','','','','',203,0,6,1,0,0,0,'',8.20,0,0.00,1,25,2,0,'',0,0,0.000,'','-',0,'','-',0,'5171.png','','',0,0.000,0),(11798,5204,'POLLO PRIMERA','POLLO PRIMERA','','','','','','',205,0,1,1,0,0,0,'',2.69,0,0.00,1,1,2,0,'',0,0,0.000,'','-',0,'','-',0,'5204.png','','',0,0.000,0),(11799,5209,'FILETES DE POLLO','FILETES DE POLLO','','','','','','',205,0,2,1,0,0,0,'',6.98,0,0.00,1,2,2,0,'',0,0,0.000,'','-',0,'','-',0,'5209.png','','',0,0.000,0),(11800,5222,'CUARTOS TRASEROS POLLO','CUARTOS TRASEROS POLLO','','','','','','',205,0,3,1,0,0,0,'',2.79,0,0.00,1,3,2,0,'',0,0,0.000,'','-',0,'','-',0,'5222.png','','',0,0.000,0),(11801,5262,'TOMATE RAMA','TOMATE RAMA','','Origen:España;Categoria:i;Calibre:m;Variedad:racimo','','','','',502,0,24,1,0,0,0,'',1.99,0,0.00,1,92,4,0,'',0,0,0.000,'','-',0,'','-',0,'5262.png','','',0,0.000,0),(11802,5269,'CEBOLLA TIERNA  MANOJO','CEBOLLA TIERNA  MANOJO','','','','','','',502,0,23,0,0,0,0,'',1.15,0,0.00,1,262,4,0,'',0,0,0.000,'','-',0,'','-',0,'5269.png','','',0,0.000,0),(11803,5272,'PAN INTEGRAL BAGUETTINA','PAN INTEGRAL BAGUETTINA','','Ingredientes:HARINA DE TRIGO,agua,SALVADO DE TRIGO,levadura,sal,citrato de potasio;Emulgentes:E-472;Puede contener trazas de:LECHE,SOJA,SESAMO','','','','',601,0,1,0,0,0,0,'',0.25,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'5272.png','','',0,0.000,0),(11804,5612,'MANOS DE CERDO','MANOS DE CERDO','','','','','','',202,0,9,1,0,0,0,'',3.99,0,0.00,1,53,2,0,'',0,0,0.000,'','-',0,'','-',0,'5612.png','','',0,0.000,0),(11805,5969,'COSTILLA+PANCETA','COSTILLA+PANCETA','','','','','','',202,0,13,1,0,0,0,'',5.98,0,0.00,1,194,2,0,'',0,0,0.000,'','-',0,'','-',0,'5969.png','','',0,0.000,0),(11806,6013,'PATE  CAMPAÑA','PATE  CAMPAÑA','','','','','','',301,0,9,1,0,0,0,'',9.55,0,0.00,1,95,3,0,'',10,0,0.000,'','-',0,'','-',0,'6013.png','','',0,0.000,0),(11807,6030,'PICADA MIXTA','PICADA MIXTA','','','','','','',208,0,2,1,0,0,0,'',6.98,0,0.00,1,210,2,0,'',3,0,0.000,'','-',0,'','-',0,'6030.png','','',0,0.000,0),(11808,6447,'CABEZA LOMO SIN HUESO','CABEZA LOMO SIN HUESO','','','','','','',202,0,13,1,0,0,0,'',6.99,0,0.00,1,38,2,0,'',0,0,0.000,'','-',0,'','-',0,'6447.png','','',0,0.000,0),(11809,6820,'PIÑA CORTADA','PIÑA CORTADA','','','','','','',501,0,17,0,0,0,0,'',1.29,0,0.00,1,0,4,0,'',3,0,0.000,'','-',0,'','-',0,'6820.png','','',0,0.000,0),(11810,7279,'ALAS DE POLLO','ALAS DE POLLO','','','','','','',205,0,4,1,0,0,0,'',3.99,0,0.00,1,4,2,0,'',0,0,0.000,'','-',0,'','-',0,'7279.png','','',0,0.000,0),(11811,7428,'ESPINAZOS POLLO GRANEL','ESPINAZOS POLLO GRANEL','','','','','','',205,0,7,1,0,0,0,'',1.39,0,0.00,1,10,2,0,'',0,0,0.000,'','-',0,'','-',0,'7428.png','','',0,0.000,0),(11812,7781,'CONEJOS','CONEJOS','','','','','','',205,0,6,1,0,0,0,'',6.25,0,0.00,1,12,2,0,'',0,0,0.000,'','-',0,'','-',0,'7781.png','','',0,0.000,0),(11813,8427,'QUESO MANCHEGO BARRA TIER','QUESO MANCHEGO BARRA TIER','','','','','','',302,0,3,1,0,0,0,'',11.75,0,0.00,1,50,3,0,'',0,0,0.000,'','-',0,'','-',0,'8427.png','','',0,0.000,0),(11814,9317,'CHORIZO CASERO','CHORIZO CASERO','','SIN GLUTEN','','','','',210,0,4,1,0,0,0,'',6.98,0,0.00,1,269,2,0,'',0,0,0.000,'','-',0,'','-',0,'9317.png','','',0,0.000,0),(11815,9446,'PAN BARRA GALLEGA','PAN BARRA GALLEGA','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura;Emulgentes:E-472(e)','','','','',601,0,2,0,0,0,0,'',0.80,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'9446.png','','',0,0.000,0),(11816,9525,'MORCILLA DE CEBOLLA','MORCILLA DE CEBOLLA','','','','','','',210,0,6,1,0,0,0,'',6.50,0,0.00,1,644,2,0,'',0,0,0.000,'','-',0,'','-',0,'9525.png','','',0,0.000,0),(11817,9649,'SOLOMILLO VACUNO','SOLOMILLO VACUNO','','','','','','',201,0,1,1,0,0,0,'',31.98,0,0.00,1,75,2,0,'',0,0,0.000,'','-',0,'','-',0,'9649.png','','',0,0.000,0),(11818,9679,'KIWI BANDEJA','KIWI BANDEJA','','','','','','',501,0,18,0,0,0,0,'',2.49,0,0.00,1,269,4,0,'',0,0,0.000,'','-',0,'','-',0,'9679.png','','',0,0.000,0),(11819,9963,'PECHUGA POLLO S/H BANDEJA','PECHUGA POLLO S/H BANDEJA','','','','','','',205,0,12,1,0,0,0,'',6.98,0,0.00,1,176,2,0,'',0,0,0.000,'','-',0,'','-',0,'9963.png','','',0,0.000,0),(11820,9972,'ALAS POLLO BANDEJA','ALAS POLLO BANDEJA','','','','','','',205,0,9,1,0,0,0,'',3.99,0,0.00,1,162,2,0,'',0,0,0.000,'','-',0,'','-',0,'9972.png','','',0,0.000,0),(11821,9975,'CUARTOS TRASEROS','CUARTOS TRASEROS','','','','','','',205,0,11,1,0,0,0,'',2.85,0,0.00,1,171,2,0,'',0,0,0.000,'','-',0,'','-',0,'9975.png','','',0,0.000,0),(11822,9978,'JAMONCITOS POLLO BANDEJA','JAMONCITOS POLLO BANDEJA','','','','','','',205,0,13,1,0,0,0,'',3.99,0,0.00,1,179,2,0,'',0,0,0.000,'','-',0,'','-',0,'9978.png','','',0,0.000,0),(11823,10060,'PAN BARRA INTEGRAL','PAN BARRA INTEGRAL','','Ingredientes:HARINA DE TRIGO,agua,SALVADO DE TRIGO,sal,levadura;Emulgentes:E-472e;Puede contener trazas de:LECHE Y DERIVADOS,SOJA Y DERIVADOS,SESAMO,FRUTOS DE CÁSCARA','','','','',601,0,3,0,0,0,0,'',0.65,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'10060.png','','',0,0.000,0),(11824,10098,'FIAMBRE DE LOMO PLANO','FIAMBRE DE LOMO PLANO','','Contiene:SOJA Y DERIVADOS','','','','',208,0,6,1,0,0,0,'',7.90,0,0.00,1,602,2,0,'',0,0,0.000,'','-',0,'','-',0,'10098.png','','',0,0.000,0),(11825,10241,'SALCHICHON IBERICO GUIJUE','SALCHICHON IBERICO GUIJUE','','Ingredientes:carne de cerdo ibérico,grasa de cerdo ib.,sal,pimienta blanca,pimienta,nuez moscada,ajo,azucar,vino blanco,LACTOSA,PROTEINAS LACTEAS,dextrosa;Conservadores:E-250,E-252;Antioxidante:E-301;Condiciones de conservacion:conservar en lugar fresco y seco','','','','',301,0,16,1,0,0,0,'',17.95,0,0.00,1,593,3,0,'',0,0,0.000,'','-',0,'','-',0,'10241.png','','',0,0.000,0),(11826,10242,'CHORIZO IBERICO GUIJUELO','CHORIZO IBERICO GUIJUELO','','Ingredientes:carne de cerdo ibérico,grasa de cerdo ib.,sal,pimenton,ajo,oregano,LACTOSA,dextrosa,PROTEINA DE SOJA;Emulgentes:E-450,E-451;Conservadores:E-250,E-252;Antioxidante:E-330,E-331;Potenciador del sabor:E-621;Condiciones de conservacion:conservar en lugar fresco y seco','','','','',301,0,17,1,0,0,0,'',17.95,0,0.00,1,594,3,0,'',0,0,0.000,'','-',0,'','-',0,'10242.png','','',0,0.000,0),(11827,10277,'QUESO BRIE (3KG.)','QUESO BRIE (3KG.)','','','','','','',302,0,14,1,0,0,0,'',9.50,0,0.00,1,597,3,0,'',0,0,0.000,'','-',0,'','-',0,'10277.png','','',0,0.000,0),(11828,10561,'CORDERO NACIONAL MEDIO/CU','CORDERO NACIONAL MEDIO/CU','','','','','','',203,0,8,1,0,0,0,'',10.50,0,0.00,1,202,2,0,'',0,0,0.000,'','-',0,'','-',0,'10561.png','','',0,0.000,0),(11829,10612,'CONTRAMUSLOS POLLO BANDEJ','CONTRAMUSLOS POLLO BANDEJ','','','','','','',205,0,10,1,0,0,0,'',4.65,0,0.00,1,169,2,0,'',0,0,0.000,'','-',0,'','-',0,'10612.png','','',0,0.000,0),(11830,10688,'HIGADO Y CORAZON POLLO','HIGADO Y CORAZON POLLO','','','','','','',205,0,14,1,0,0,0,'',3.30,0,0.00,1,180,2,0,'',0,0,0.000,'','-',0,'','-',0,'10688.png','','',0,0.000,0),(11831,10830,'QUESO MEZCLA SEMI','QUESO MEZCLA SEMI','','','','','','',302,0,1,1,0,0,0,'',11.99,0,0.00,1,5,3,0,'',0,0,0.000,'','-',0,'','-',0,'10830.png','','',0,0.000,0),(11832,11509,'MORTADELA DE PAVO','MORTADELA DE PAVO','','','','','','',301,0,18,1,0,0,0,'',6.95,0,0.00,1,648,3,0,'',30,0,0.000,'','-',0,'','-',0,'11509.png','','',0,0.000,0),(11833,11617,'FIAMBRE BARRA SANDWICH BA','FIAMBRE BARRA SANDWICH BA','','','','','','',301,0,1,1,0,0,0,'',5.95,0,0.00,1,7,3,0,'',30,0,0.000,'','-',0,'','-',0,'11617.png','','',0,0.000,0),(11834,11788,'FILETES PECHUGA PAVO BAND','FILETES PECHUGA PAVO BAND','','','','','','',205,0,17,1,0,0,0,'',8.50,0,0.00,1,664,2,0,'',0,0,0.000,'','-',0,'','-',0,'11788.png','','',0,0.000,0),(11835,11862,'PATE PIMIENTA VERDE','PATE PIMIENTA VERDE','','','','','','',301,0,19,1,0,0,0,'',13.95,0,0.00,1,663,3,0,'',10,0,0.000,'','-',0,'','-',0,'11862.png','','',0,0.000,0),(11836,11865,'PATE DE MANZANA','PATE DE MANZANA','','','','','','',301,0,20,1,0,0,0,'',17.90,0,0.00,1,664,3,0,'',10,0,0.000,'','-',0,'','-',0,'11865.png','','',0,0.000,0),(11837,12075,'LOTE OFERTA CANICERIA - 1','LOTE OFERTA CANICERIA - 1','','','','','','',208,0,4,1,0,0,0,'',8.00,0,0.00,1,665,2,0,'',0,0,0.000,'','-',0,'','-',0,'12075.png','','',0,0.000,0),(11838,12470,'LONGANIZA BLANCA BANDEJA','LONGANIZA BLANCA BANDEJA','','Ingredientes:carne de cerdo,agua,fibra vegetal,almidon,sal,maltodextrina,dextrosa,especias;Conservadores:E-221 (SULFITOS),aromas;Antioxidantes:E-300,E-331;Colorantes:E-120;Contiene:SULFITOS','','','','',210,0,7,1,0,0,0,'',6.98,0,0.00,1,678,2,0,'',0,0,0.000,'','-',0,'','-',0,'12470.png','','',0,0.000,0),(11839,12490,'CHORIZO CASERO BANDEJA','CHORIZO CASERO BANDEJA','','Ingredientes:carne de cerdo,agua,fibra vegetal,almidon,sal,pimenton,dextrina;Conservadores:E-262,maltodextrina de maiz,especias;Antioxidantes:E-331,E-301,aroma;Colorantes:E-120','','','','',210,0,8,1,0,0,0,'',6.98,0,0.00,1,679,2,0,'',0,0,0.000,'','-',0,'','-',0,'12490.png','','',0,0.000,0),(11840,12700,'PAN FIBRA VERDE 170GR','PAN FIBRA VERDE 170GR','','Ingredientes:HARINA DE TRIGO,agua;Fibras:SALVADO DE TRIGO,espinacas,lino,estabilizante,E-410,emulgentes,E-472e,sal,levadura,GLUTEN DE TRIGO;Puede contener trazas de:SESAMO','','','','',601,0,4,0,0,0,0,'',0.75,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'12700.png','','',0,0.000,0),(11841,12922,'JUDIA ANCHA BANDEJA','JUDIA ANCHA BANDEJA','','','','','','',502,0,24,0,0,0,0,'',2.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'12922.png','','',0,0.000,0),(11842,13871,'PAN BAGUETTINA','PAN BAGUETTINA','','Ingredientes:HARINA DE TRIGO,agua,saborizantes,levadura;Emulgentes:E-472(e);Puede contener trazas de:SOJA,SESAMO','','','','',601,0,5,0,0,0,0,'',0.22,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'13871.png','','',0,0.000,0),(11843,13880,'CROISSANT ARTESANITO','CROISSANT ARTESANITO','','Ingredientes:HARINA DE TRIGO;Margarina:grasa de palma,aceite de soja,aceite de girasol,jarabe de glucosa-fructosa,agua;Emulgentes:E-322,E-471;Acidulantes:E-330,aroma;Colorantes:E-160a,agua,levadura,azucar,sal,GLUTEN;Emulgentes:E-471,HUEVO;Puede contener trazas de:SOJA Y DERIVADOS,LECHE Y DERIVADOS','','','','',602,0,1,0,0,0,0,'',5.75,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'13880.png','','',0,0.000,0),(11844,13922,'SNEECKEN PASAS-NUECES UNI','SNEECKEN PASAS-NUECES UNI','','Ingredientes:Masa:HARINA DE TRIGO;Margarina:agua,azucar,levadura,sal,vainilla,GLUTEN;Emulgentes:E-471;Relleno:Crema pastelera:Almidón modificado:E-1414,SUERO LACTEO,dextrosa,grasa de palma,grasa de coco,ALBUMINA DE HUEVO;Estabilizantes:E-339,E-450,E-516,jarabe de glucosa,sal,aroma;Espesantes:E-401;Acidificantes:E-504,PROTEINAS LACTEAS;Colorantes:E-170,E-160a,extracto de paprika,pasas,NUECES,HUEVO,estabilizador,E-440,E-415,conservantes,E-202,guindas;Puede contener trazas de:PESCADO Y DERIVADOS,SOJA Y DERIVADOS,SESAMO','','','','',602,0,2,0,0,0,0,'',0.80,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'13922.png','','',0,0.000,0),(11845,13985,'PASTISSET BONIATO','PASTISSET BONIATO','','Ingredientes:confitura de boniato,(,boniato,azucar,dextrosa;Espesantes:E-412,conservador,E-202,aroma de limon;Antioxidante:E-300,),HARINA DE TRIGO,manteca de cerdo,agua,azucar,EXTRACTO DE Malta,aroma de anis;Puede contener trazas de:PESCADO,LECHE,FRUTOS DE CÁSCARA,SOJA,HUEVO','','','','',602,0,3,0,0,0,0,'',8.45,0,0.00,1,0,6,0,'',7,0,0.000,'','-',0,'','-',0,'13985.png','','',0,0.000,0),(11846,14032,'MOUSSE SUPREMA 3 SABORES','MOUSSE SUPREMA 3 SABORES','','','','','','',301,0,21,1,0,0,0,'',19.95,0,0.00,1,720,3,0,'',10,0,0.000,'','-',0,'','-',0,'14032.png','','',0,0.000,0),(11847,14058,'EMBUTIDO SURTIDO BARBACOA','EMBUTIDO SURTIDO BARBACOA','','Ingredientes:carne de cerdo,agua,fibra vegetal,almidon,QUESO,espinacas,sal,pimenton,maltodextrina de maiz,dextrosa,jarabe de glucosa,especias;Conservadores:E-221,sultitos,aroma;Antioxidantes:E-300,E-331;Estabilizantes:E-450(i),E-451;Colorante natural:E-120','','','','',210,0,9,1,0,0,0,'',7.95,0,0.00,1,730,2,0,'',0,0,0.000,'','-',0,'','-',0,'14058.png','','',0,0.000,0),(11848,14059,'LONGANIZA DE POLLO BANDEJ','LONGANIZA DE POLLO BANDEJ','','','','','','',210,0,10,1,0,0,0,'',7.25,0,0.00,1,731,2,0,'',0,0,0.000,'','-',0,'','-',0,'14059.png','','',0,0.000,0),(11849,14106,'MAIZ MAZORCA COCIDO 2U.','MAIZ MAZORCA COCIDO 2U.','','','','','','',502,0,25,0,0,0,0,'',1.19,0,0.00,1,147,4,0,'',0,0,0.000,'','-',0,'','-',0,'14106.png','','',0,0.000,0),(11850,14107,'COL LOMBARDA UNIDAD','COL LOMBARDA UNIDAD','','Origen:España;Categoria:1º;Variedad:rojo;Calibre:8','','','','',502,0,26,0,0,0,0,'',1.95,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'14107.png','','',0,0.000,0),(11851,14108,'ENDIBIA BANDEJA','ENDIBIA BANDEJA','','','','','','',502,0,27,0,0,0,0,'',1.59,0,0.00,1,116,4,0,'',0,0,0.000,'','-',0,'','-',0,'14108.png','','',0,0.000,0),(11852,14110,'LECHUGA ICEBERG','LECHUGA ICEBERG','','','','','','',502,0,28,0,0,0,0,'',0.99,0,0.00,1,146,4,0,'',0,0,0.000,'','-',0,'','-',0,'14110.png','','',0,0.000,0),(11853,14111,'PIMIENTOS PADRON BANDEJA','PIMIENTOS PADRON BANDEJA','','','','','','',502,0,29,0,0,0,0,'',1.39,0,0.00,1,129,4,0,'',0,0,0.000,'','-',0,'','-',0,'14111.png','','',0,0.000,0),(11854,14112,'REMOLACHA PRECOCIDA BANDE','REMOLACHA PRECOCIDA BANDE','','','','','','',502,0,30,0,0,0,0,'',0.99,0,0.00,1,125,4,0,'',0,0,0.000,'','-',0,'','-',0,'14112.png','','',0,0.000,0),(11855,14223,'PAN BOCADILLOS PEQUEÑOS','PAN BOCADILLOS PEQUEÑOS','','','','','','',601,0,6,0,0,0,0,'',0.45,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'14223.png','','',0,0.000,0),(11856,14224,'PAN BOCADILLOS GRANDES','PAN BOCADILLOS GRANDES','','','','','','',601,0,7,0,0,0,0,'',0.55,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'14224.png','','',0,0.000,0),(11857,14225,'PAN LINIETES','PAN LINIETES','','','','','','',601,0,8,0,0,0,0,'',0.55,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'14225.png','','',0,0.000,0),(11858,14227,'PAN BARRAS','PAN BARRAS','','','','','','',601,0,9,0,0,0,0,'',0.80,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'14227.png','','',0,0.000,0),(11859,14228,'PAN','PAN','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura,HARINA DE CENTENO,mejorante panario autorizado;Emulgentes:E-472;Antioxidante:E-300,amilasas,peso: 1kg','','','','',601,0,10,0,0,0,0,'',2.70,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'14228.png','','',0,0.000,0),(11860,14229,'PAN','PAN','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura,HARINA DE CENTENO,mejorante panario autorizado;Emulgentes:E-472;Antioxidante:E-300,amilasas,peso: 500g','','','','',601,0,11,0,0,0,0,'',1.50,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'14229.png','','',0,0.000,0),(11861,15105,'BROCHETAS DE POLLO GRANEL','BROCHETAS DE POLLO GRANEL','','','','','','',208,0,5,1,0,0,0,'',9.90,0,0.00,1,433,2,0,'',3,0,0.000,'','-',0,'','-',0,'15105.png','','',0,0.000,0),(11862,15115,'OSOBUCCO VACUNO','OSOBUCCO VACUNO','','','','','','',201,0,4,1,0,0,0,'',7.98,0,0.00,1,95,2,0,'',0,0,0.000,'','-',0,'','-',0,'15115.png','','',0,0.000,0),(11863,15234,'CANILLA VACUNO','CANILLA VACUNO','','','','','','',201,0,9,1,0,0,0,'',3.30,0,0.00,1,100,2,0,'',0,0,0.000,'','-',0,'','-',0,'15234.png','','',0,0.000,0),(11864,15247,'CEBOLLA FRANCESA MALLA','CEBOLLA FRANCESA MALLA','','','','','','',502,0,31,0,0,0,0,'',1.30,0,0.00,1,134,4,0,'',0,0,0.000,'','-',0,'','-',0,'15247.png','','',0,0.000,0),(11865,15693,'COL RIZADA 1,5KG','COL RIZADA 1,5KG','','Origen:España;Categoria:1º;Variedad:rizada;Calibre:6','','','','',502,0,32,0,0,0,0,'',1.80,0,0.00,1,140,4,0,'',0,0,0.000,'','-',0,'','-',0,'15693.png','','',0,0.000,0),(11866,15705,'LECHUGA ROMANA','LECHUGA ROMANA','','','','','','',502,0,33,0,0,0,0,'',1.15,0,0.00,1,118,4,0,'',0,0,0.000,'','-',0,'','-',0,'15705.png','','',0,0.000,0),(11867,15712,'PATATA BOLSA 5 KG.','PATATA BOLSA 5 KG.','','','','','','',502,0,34,0,0,0,0,'',4.75,0,0.00,1,123,4,0,'',0,0,0.000,'','-',0,'','-',0,'15712.png','','',0,0.000,0),(11868,15714,'PATATA ROJA 2 KG.','PATATA ROJA 2 KG.','','','','','','',502,0,35,0,0,0,0,'',2.95,0,0.00,1,152,4,0,'',0,0,0.000,'','-',0,'','-',0,'15714.png','','',0,0.000,0),(11869,15715,'PATATA LAVADA','PATATA LAVADA','','Origen:España;Variedad:mona lisa;Categoria:primera;Calibre:40','','','','',502,0,2,1,0,0,0,'',1.49,0,0.00,1,6,4,0,'',0,0,0.000,'','-',0,'','-',0,'15715.png','','',0,0.000,0),(11870,15719,'PEREJIL','PEREJIL','','','','','','',502,0,37,0,0,0,0,'',0.99,0,0.00,1,156,4,0,'',0,0,0.000,'','-',0,'','-',0,'15719.png','','',0,0.000,0),(11871,15720,'CHAMPIÑON LAMINADO BANDEJ','CHAMPIÑON LAMINADO BANDEJ','','','','','','',502,0,38,0,0,0,0,'',1.45,0,0.00,1,168,4,0,'',0,0,0.000,'','-',0,'','-',0,'15720.png','','',0,0.000,0),(11872,15722,'RAIZ DE APIO','RAIZ DE APIO','','Origen:holanda;Categoria:i;Calibre:inf. 1 kg','','','','',502,0,23,1,0,0,0,'',2.80,0,0.00,1,89,4,0,'',0,0,0.000,'','-',0,'','-',0,'15722.png','','',0,0.000,0),(11873,15728,'TOMATE ENSALADA','TOMATE ENSALADA','','Origen:España;Categoria:i;Calibre:g;Variedad:caramba','','','','',502,0,10,1,0,0,0,'',1.85,0,0.00,1,30,4,0,'',0,0,0.000,'','-',0,'','-',0,'15728.png','','',0,0.000,0),(11874,15729,'TOMATE CHERRY','TOMATE CHERRY','','','','','','',502,0,41,0,0,0,0,'',0.99,0,0.00,1,159,4,0,'',0,0,0.000,'','-',0,'','-',0,'15729.png','','',0,0.000,0),(11875,15731,'TOMATE RAF','TOMATE RAF','','Origen:España;Categoria:i;Variedad:raff;Calibre:57/82','','','','',502,0,1,1,0,0,0,'',2.75,0,0.00,1,4,4,0,'',0,0,0.000,'','-',0,'','-',0,'15731.png','','',0,0.000,0),(11876,15733,'CASTAÑAS','CASTAÑAS','','Origen:España;Categoria:i;Calibre:80/100;Variedad:ver envase','','','','',505,0,1,1,0,0,0,'',5.49,0,0.00,1,19,4,0,'',0,0,0.000,'','-',0,'','-',0,'15733.png','','',0,0.000,0),(11877,15743,'MANGOS','MANGOS','','Origen:España;Categoria:i;Variedad:osteen;Calibre:7','','','','',501,0,23,1,0,0,0,'',2.99,0,0.00,1,74,4,0,'',0,0,0.000,'','-',0,'','-',0,'15743.png','','',0,0.000,0),(11878,15745,'TOMATE PERA','TOMATE PERA','','Origen:España;Categoria:i;Calibre:47/57;Variedad:ruth/baobab','','','','',502,0,11,1,0,0,0,'',1.69,0,0.00,1,41,4,0,'',0,0,0.000,'','-',0,'','-',0,'15745.png','','',0,0.000,0),(11879,15747,'MELON PIEL DE SAPO EXTRA','MELON PIEL DE SAPO EXTRA','','Origen:Brasil;Categoria:i;Calibre:5;Variedad:piel de sapo','','','','',501,0,17,1,0,0,0,'',1.99,0,0.00,1,58,4,0,'',0,0,0.000,'','-',0,'','-',0,'15747.png','','',0,0.000,0),(11880,15753,'ZANAHORIA BOLSA  1 KG.','ZANAHORIA BOLSA  1 KG.','','','','','','',502,0,44,0,0,0,0,'',0.79,0,0.00,1,133,4,0,'',0,0,0.000,'','-',0,'','-',0,'15753.png','','',0,0.000,0),(11881,15776,'POLLO ENTERO CATEG.A BJA','POLLO ENTERO CATEG.A BJA','','','','','','',205,0,15,1,0,0,0,'',2.44,0,0.00,1,214,2,0,'',0,0,0.000,'','-',0,'','-',0,'15776.png','','',0,0.000,0),(11882,16160,'PASTELERIA GRANEL','PASTELERIA GRANEL','','','','','','',603,0,1,1,0,0,0,'',3.95,0,0.00,1,10,6,0,'',60,0,0.000,'','-',0,'','-',0,'16160.png','','',0,0.000,0),(11883,16628,'PALETILLA CORDERO ENTERA','PALETILLA CORDERO ENTERA','','','','','','',203,0,8,1,0,0,0,'',12.99,0,0.00,1,24,2,0,'',0,0,0.000,'','-',0,'','-',0,'16628.png','','',0,0.000,0),(11884,17107,'BURGER DE VACUNO Y CERDO','BURGER DE VACUNO Y CERDO','','','','','','',201,0,10,1,0,0,0,'',6.98,0,0.00,1,276,2,0,'',3,0,0.000,'','-',0,'','-',0,'17107.png','','',0,0.000,0),(11885,17232,'CHORIZO PAMPLONES','CHORIZO PAMPLONES','','Ingredientes:carne de cerdo,tocino de cerdo,sal,pimenton,PROTEINA DE SOJA,LECHE EN POLVO,LACTOSA,dextrina,azucar,especias,dextrosa;Emulgentes:E-450(i);Antioxidante:E-301;Conservadores:E-250,E-252;Colorantes:E-120;Condiciones de conservacion:conservar en lugar fresco y seco','','','','',301,0,8,1,0,0,0,'',9.99,0,0.00,1,65,3,0,'',0,0,0.000,'','-',0,'','-',0,'17232.png','','',0,0.000,0),(11886,17262,'CARETAS CERDO','CARETAS CERDO','','','','','','',202,0,8,1,0,0,0,'',3.99,0,0.00,1,52,2,0,'',0,0,0.000,'','-',0,'','-',0,'17262.png','','',0,0.000,0),(11887,17418,'ARANDANOS','ARANDANOS','','','','','','',501,0,21,0,0,0,0,'',2.99,0,0.00,1,148,4,0,'',0,0,0.000,'','-',0,'','-',0,'17418.png','','',0,0.000,0),(11888,17725,'ENSAIMADA MINI HOJALDRADA','ENSAIMADA MINI HOJALDRADA','','Ingredientes:HARINA DE TRIGO,grasa de cerdo;Estabilizadores:E-471,azucar,agua,azucar,levadura,HUEVO,sal,dextrosa,antiapelmazante,E-170;Puede contener trazas de:APIO,SOJA Y DERIVADOS,LECHE Y DERIVADOS,SESAMO,PESCADO Y DERIVADOS,FRUTOS DE CÁSCARA','','','','',602,0,4,0,0,0,0,'',5.24,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'17725.png','','',0,0.000,0),(11889,17863,'PIERNA CORDERO ENTERA','PIERNA CORDERO ENTERA','','','','','','',203,0,9,1,0,0,0,'',12.99,0,0.00,1,23,2,0,'',0,0,0.000,'','-',0,'','-',0,'17863.png','','',0,0.000,0),(11890,17864,'CHULETAS CORDERO PALO + R','CHULETAS CORDERO PALO + R','','','','','','',203,0,10,1,0,0,0,'',16.95,0,0.00,1,20,2,0,'',0,0,0.000,'','-',0,'','-',0,'17864.png','','',0,0.000,0),(11891,18308,'PECHUGA POLLO FILETEADA B','PECHUGA POLLO FILETEADA B','','','','','','',205,0,16,1,0,0,0,'',7.95,0,0.00,1,420,2,0,'',0,0,0.000,'','-',0,'','-',0,'18308.png','','',0,0.000,0),(11892,18310,'GAMBON CAJA 2KG ALTA MAR','GAMBON CAJA 2KG ALTA MAR','','','','','','',702,0,8,1,0,0,0,'',21.80,0,0.00,1,141,5,0,'',0,0,0.000,'','-',0,'','-',0,'18310.png','','',0,0.000,0),(11893,18402,'LONGANIZA BLANCA EXTRA','LONGANIZA BLANCA EXTRA','','Contiene:SULFITOS,SIN GLUTEN','','','','',210,0,2,1,0,0,0,'',6.98,0,0.00,1,267,2,0,'',0,0,0.000,'','-',0,'','-',0,'18402.png','','',0,0.000,0),(11894,18404,'LONGANIZA ROJA','LONGANIZA ROJA','','Contiene:SULFITOS','','','','',210,0,3,1,0,0,0,'',6.98,0,0.00,1,268,2,0,'',0,0,0.000,'','-',0,'','-',0,'18404.png','','',0,0.000,0),(11895,18407,'CHORIZO MINI PICANTE','CHORIZO MINI PICANTE','','Ingredientes:carne de cerdo,agua,fibra vegetal,almidon,sal,pimenton,dextrina;Conservadores:E-262,maltodextrina de maiz,especias,cayena;Antioxidante:E-331,E-301,aromas;Colorantes:E-120','','','','',210,0,5,1,0,0,0,'',6.98,0,0.00,1,271,2,0,'',0,0,0.000,'','-',0,'','-',0,'18407.png','','',0,0.000,0),(11896,18463,'PAN BARRA','PAN BARRA','','Ingredientes:HARINA DE TRIGO,agua,levadura,sal;Emulgentes:E-472(e);Puede contener trazas de:SOJA,LECHE,SESAMO','','','','',601,0,12,0,0,0,0,'',0.45,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'18463.png','','',0,0.000,0),(11897,18595,'JAMON DE HEMBRA DESH. RES','JAMON DE HEMBRA DESH. RES','','Ingredientes:jamon de cerdo,sal,azucar;Conservadores:E-250,E-252;Antioxidante:E-302;Condiciones de conservacion:conservar en lugar fresco y seco','','','','',301,0,2,1,0,0,0,'',18.95,0,0.00,1,9,3,0,'',0,0,0.000,'','-',0,'','-',0,'18595.png','','',0,0.000,0),(11898,18730,'SOBRASADA MARINA CERDO DU','SOBRASADA MARINA CERDO DU','','','','','','',208,0,2,1,0,0,0,'',7.80,0,0.00,1,272,2,0,'',0,0,0.000,'','-',0,'','-',0,'18730.png','','',0,0.000,0),(11899,18969,'ALAS PARTIDAS BANDEJA','ALAS PARTIDAS BANDEJA','','','','','','',205,0,18,1,0,0,0,'',5.35,0,0.00,1,939,2,0,'',0,0,0.000,'','-',0,'','-',0,'18969.png','','',0,0.000,0),(11900,19060,'SECRETO DE CERDO CRUCE DU','SECRETO DE CERDO CRUCE DU','','','','','','',202,0,16,1,0,0,0,'',6.98,0,0.00,1,523,2,0,'',0,0,0.000,'','-',0,'','-',0,'19060.png','','',0,0.000,0),(11901,19077,'PAN DE LECHE','PAN DE LECHE','','Ingredientes:HARINA DE TRIGO,agua,LECHE DES. EN POLVO;Margarina:aceite de palma,aceite de coco,aceite de girasol,agua,sal;Emulgentes:E-471,lecitina de girasol,aromas;Acidulantes:acido citrico;Conservadores:E-202;Colorante natural:E-160a,levadura,azucar,sal,HARINA DE Malta DE TRIGO,aroma;Puede contener trazas de:SESAMO,SOJA','','','','',601,0,13,0,0,0,0,'',0.33,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'19077.png','','',0,0.000,0),(11902,19421,'QUESO MAASDAM','QUESO MAASDAM','','','','','','',302,0,13,1,0,0,0,'',7.95,0,0.00,1,303,3,0,'',0,0,0.000,'','-',0,'','-',0,'19421.png','','',0,0.000,0),(11903,19717,'QUESO CAMEMBERT GIGANTE','QUESO CAMEMBERT GIGANTE','','','','','','',302,0,5,1,0,0,0,'',8.75,0,0.00,1,98,3,0,'',0,0,0.000,'','-',0,'','-',0,'19717.png','','',0,0.000,0),(11904,20115,'FILETE DE BACALADILLA','FILETE DE BACALADILLA','','','','','','',701,0,8,1,0,0,0,'',6.80,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'20115.png','','',0,0.000,0),(11905,20120,'CAMARON ROJO','CAMARON ROJO','','','','','','',702,0,9,1,0,0,0,'',10.90,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'20120.png','','',0,0.000,0),(11906,20125,'TIRAS DE POTON','TIRAS DE POTON','','','','','','',702,0,5,1,0,0,0,'',6.60,0,0.00,1,115,5,0,'',0,0,0.000,'','-',0,'','-',0,'20125.png','','',0,0.000,0),(11907,20134,'LANGOSTINO COCIDO 30/40 G','LANGOSTINO COCIDO 30/40 G','','','','','','',702,0,7,1,0,0,0,'',13.95,0,0.00,1,126,5,0,'',0,0,0.000,'','-',0,'','-',0,'20134.png','','',0,0.000,0),(11908,20289,'MORTADELA SICILIANA','MORTADELA SICILIANA','','','','','','',301,0,12,1,0,0,0,'',5.45,0,0.00,1,266,3,0,'',30,0,0.000,'','-',0,'','-',0,'20289.png','','',0,0.000,0),(11909,21506,'CARACOLA CHOCOLATE MINI','CARACOLA CHOCOLATE MINI','','Ingredientes:Masa:HARINA DE TRIGO,agua,grasa de palma,aceite de soja y girasol,emulgentes,E-471,sal,acidulante,E-330,aromas,antioxidantes,E-306,E-304,azucar,GLUTEN,levadura;Relleno:grasa de palma,aceite de girasol,cacao en polvo desgrasado,almidon de maiz,almidon de patata,PASTA DE AVELLANAS;Emulgentes:E-322,aromas;Cobertura:HUEVO,gelatina,fideo chocolate;Puede contener trazas de:LECHE Y DERIVADOS,SOJA Y DERIVADOS','','','','',602,0,5,0,0,0,0,'',5.26,0,0.00,1,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'21506.png','','',0,0.000,0),(11910,21508,'CROISSANT ARTESANO','CROISSANT ARTESANO','','Ingredientes:HARINA DE TRIGO;Margarina:grasa de palma,agua,aceite de girasol,aceite de soja;Emulgentes:E-471,sal,HUEVO;Acidificantes:E-330,aromas;Colorantes:E-160a,agua,levadura,azucar,sal,GLUTEN;Emulgentes:E-471;Puede contener trazas de:SOJA,LECHE','','','','',602,0,6,0,0,0,0,'',0.60,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'21508.png','','',0,0.000,0),(11911,21510,'NAPOLITANA CHOCOLATE','NAPOLITANA CHOCOLATE','','Ingredientes:Masa:HARINA DE TRIGO;Margarina:grasa de palma,agua,aceite de soja,aceite de girasol;Emulgentes:E-471,sal;Acidulantes:E-330;Antioxidante:E-306,E-304,aroma;Colorantes:E-160a,agua,azucar,levadura,sal,GLUTEN DE TRIGO;Emulgentes:E-471,E-472;Espesantes:E-466,FIBRA DE TRIGO;Relleno:azucar,grasa de palma,aceite de girasol,cacao en polvo desgrasado,almidon de maiz,PASTA DE AVELLANAS;Emulgentes:E-322,aromas;Cobertura:HUEVO,jarabe de glucosa,estabilizador,E-440,E-415,regulador de acidez,E-450,conservantes,E-202,talco E-553b;Puede contener trazas de:SOJA,LECHE Y DERIVADOS,FRUTOS DE CÁSCARA','','','','',602,0,7,0,0,0,0,'',0.60,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'21510.png','','',0,0.000,0),(11912,21511,'NAPOLITANA CREMA','NAPOLITANA CREMA','','Ingredientes:Masa:HARINA DE TRIGO;Margarina:grasa de palma,aceite de soja,aceite de girasol;Emulgentes:E-471,sal;Acidulantes:E-330,aromas;Antioxidante:E-306,E-304;Colorantes:E-160a,agua,azucar,levadura,GLUTEN;Emulgentes:E-472e;Relleno:Almidón modificado:E-1414,SUERO LACTEO,dextrosa;Espesantes:E-401,ALBUMINA DE HUEVO;Estabilizantes:E-339,E-450,E-516,glucosa;Acidificantes:E-504,PROTEINAS LACTEAS;Colorantes:E-170,E-160a,E-160c,HUEVO,dextrosa,antiapelmazante,E-170;Puede contener trazas de:SOJA,SESAMO,PESCADO,FRUTOS DE CÁSCARA','','','','',602,0,8,0,0,0,0,'',0.60,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'21511.png','','',0,0.000,0),(11913,21514,'ENSAIMADA','ENSAIMADA','','Ingredientes:HARINA DE TRIGO,manteca de cerdo,azucar,levadura,HUEVO,sal,dextrosa,antiapelmazante,E-170;Puede contener trazas de:APIO,SOJA,LECHE,SESAMO,FRUTOS DE CÁSCARA,PESCADO','','','','',602,0,9,0,0,0,0,'',0.70,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'21514.png','','',0,0.000,0),(11914,21518,'TARTA DE MANZANA','TARTA DE MANZANA','','Ingredientes:HARINA DE TRIGO,margarina,agua,sal,jarabe de glucosa;Emulgentes:E-472e;Antiapelmazantes:E-440,E-551;Relleno:manzana;Antioxidante:E-300,sal,agua;Crema:azucar;Almidón modificado:E-1414,SUERO DE LECHE EN POLVO,grasa de palma,dextrosa,aceite de coco;Estabilizantes:E-339,E-516,E-450;Espesantes:E-401;Acidulantes:E-504,E-330,ALBUMINA DE HUEVO,aroma,PROTEINAS LACTEAS,LACTOSA;Colorantes:E-160a,E-170,E-127,extracto de paprika,estabilizador,E-440,E-415,conservantes,E-202,CEREZA CONFITADA (SULFITOS);Puede contener trazas de:SOJA Y DERIVADOS,FRUTOS DE CÁSCARA,SESAMO,APIO,PESCADO','','','','',602,0,10,0,0,0,0,'',1.25,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'21518.png','','',0,0.000,0),(11915,21585,'CROISSANT MINI RECTO','CROISSANT MINI RECTO','','Ingredientes:HARINA DE TRIGO;Margarina:grasa de palma,ACEITE DE SOJA,aceite de girasol,jarabe de glucosa-fructosa,agua;Emulgentes:E-322,E-471;Acidulantes:E-330,aroma;Colorantes:E-160a,agua,azucar,levadura,sal,GLUTEN,HUEVO;Puede contener trazas de:LECHE Y DERIVADOS,FRUTOS DE CÁSCARA','','','','',602,0,11,0,0,0,0,'',6.20,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'21585.png','','',0,0.000,0),(11916,21593,'PALMERA HOJALDRE','PALMERA HOJALDRE','','Ingredientes:Masa:HARINA DE TRIGO;Margarina:grasa de palma,ACEITE DE SOJA,aceite de girasol,jarabe de glucosa-fructosa,agua;Emulgentes:E-471;Acidulantes:E-330,aromas;Colorantes:E-160a,agua,sal,jarabe de glucosa;Emulgentes:E-472e;Estabilizantes:E-440;Antiapelmazantes:E-551;Cobertura:azucar;Decoracion:chocolate fondant;Puede contener trazas de:SOJA,HUEVO,FRUTOS DE CÁSCARA','','','','',602,0,12,0,0,0,0,'',0.60,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'21593.png','','',0,0.000,0),(11917,23020,'CHULETERO SIN SOLOMILLO','CHULETERO SIN SOLOMILLO','','','','','','',202,0,15,1,0,0,0,'',4.98,0,0.00,1,457,2,0,'',0,0,0.000,'','-',0,'','-',0,'23020.png','','',0,0.000,0),(11918,23214,'BLANQUET','BLANQUET','','Contiene:HUEVO Y DERIVADOS','','','','',210,0,1,1,0,0,0,'',6.98,0,0.00,1,126,2,0,'',0,0,0.000,'','-',0,'','-',0,'23214.png','','',0,0.000,0),(11919,23301,'SURTIDO ESPECIAL NAVIDAD','SURTIDO ESPECIAL NAVIDAD','','','','','','',603,0,2,1,0,0,0,'',1.98,0,0.00,1,15,6,0,'',0,0,0.000,'','-',0,'','-',0,'23301.png','','',0,0.000,0),(11920,23395,'FRAMBUESA','FRAMBUESA','','','','','','',501,0,22,0,0,0,0,'',2.29,0,0.00,1,228,4,0,'',0,0,0.000,'','-',0,'','-',0,'23395.png','','',0,0.000,0),(11921,23414,'LIMONES MALLA','LIMONES MALLA','','','','','','',501,0,23,0,0,0,0,'',1.49,0,0.00,1,229,4,0,'',0,0,0.000,'','-',0,'','-',0,'23414.png','','',0,0.000,0),(11922,23423,'MANZANA PARA ASAR REINETA','MANZANA PARA ASAR REINETA','','Calibre:75/80;Categoria:i;Origen:España;Variedad:reineta verde','','','','',501,0,8,1,0,0,0,'',1.79,0,0.00,1,16,4,0,'',0,0,0.000,'','-',0,'','-',0,'23423.png','','',0,0.000,0),(11923,23588,'MANZANA GOLDEN','MANZANA GOLDEN','','Origen:España;Categoria:i;Calibre:70+;Variedad:golden','','','','',501,0,11,1,0,0,0,'',1.59,0,0.00,1,44,4,0,'',0,0,0.000,'','-',0,'','-',0,'23588.png','','',0,0.000,0),(11924,23731,'PAPAYA','PAPAYA','','Origen:Brasil;Categoria:primera','','','','',501,0,30,1,0,0,0,'',4.95,0,0.00,1,96,4,0,'',0,0,0.000,'','-',0,'','-',0,'23731.png','','',0,0.000,0),(11925,23795,'PERA CONFERENCIA','PERA CONFERENCIA','','Origen:España;Categoria:i;Calibre:18/20;Variedad:conferencia','','','','',501,0,31,1,0,0,0,'',1.69,0,0.00,1,98,4,0,'',0,0,0.000,'','-',0,'','-',0,'23795.png','','',0,0.000,0),(11926,23955,'MANZANA FUJI','MANZANA FUJI','','Origen:España;Categoria:i;Calibre:75/85;Variedad:fuji','','','','',501,0,10,1,0,0,0,'',1.79,0,0.00,1,25,4,0,'',0,0,0.000,'','-',0,'','-',0,'23955.png','','',0,0.000,0),(11927,24047,'CUERPOS COCIDOS','CUERPOS COCIDOS','','','','','','',702,0,9,1,0,0,0,'',19.50,0,0.00,1,244,5,0,'',0,0,0.000,'','-',0,'','-',0,'24047.png','','',0,0.000,0),(11928,24794,'PAN 6 CEREALES','PAN 6 CEREALES','','Ingredientes:HARINA DE TRIGO,agua,HARINA DE CENTENO,TRIGO,GLUTEN DE TRIGO,TRIGO SARRACENO,linaza marron,pipas de girasol,SESAMO,harina de arroz,HARINA DE Malta DE TRIGO,MALTA DE TRIGO,lino dorado,sal,levadura;Emulgentes:E-472e;Puede contener trazas de:SOJA Y DERIVADOS,LECHE Y DERIVADOS,FRUTOS DE CÁSCARA','','','','',601,0,14,0,0,0,0,'',1.25,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'24794.png','','',0,0.000,0),(11929,25003,'OREJA DE CERDO COCIDA','OREJA DE CERDO COCIDA','','','','','','',208,0,5,1,0,0,0,'',3.39,0,0.00,1,498,2,0,'',0,0,0.000,'','-',0,'','-',0,'25003.png','','',0,0.000,0),(11930,25354,'NARANJA MALLA 3KG','NARANJA MALLA 3KG','','','','','','',501,0,29,0,0,0,0,'',3.25,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'25354.png','','',0,0.000,0),(11931,25699,'PATATA GUARNICION    1,5K','PATATA GUARNICION    1,5K','','','','','','',502,0,45,0,0,0,0,'',1.89,0,0.00,1,340,4,0,'',0,0,0.000,'','-',0,'','-',0,'25699.png','','',0,0.000,0),(11932,25990,'MERLUZA DE PINCHO','MERLUZA DE PINCHO','','','','','','',701,0,5,1,0,0,0,'',11.50,0,0.00,1,150,5,0,'',0,0,0.000,'','-',0,'','-',0,'25990.png','','',0,0.000,0),(11933,26110,'(GRANEL) CUADRADOS TRUFA','(GRANEL) CUADRADOS TRUFA','','','','','','',603,0,8,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'26110.png','','',0,0.000,0),(11934,26112,'(GRANEL) CAÑAS CABELLO  L','(GRANEL) CAÑAS CABELLO  L','','','','','','',603,0,9,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'26112.png','','',0,0.000,0),(11935,26113,'(GRANEL) PETISU CREMA','(GRANEL) PETISU CREMA','','','','','','',603,0,10,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'26113.png','','',0,0.000,0),(11936,26115,'(GRANEL) MIGUELITOS CREMA','(GRANEL) MIGUELITOS CREMA','','','','','','',603,0,11,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'26115.png','','',0,0.000,0),(11937,26448,'QUESO PURO OVEJA CURADO R','QUESO PURO OVEJA CURADO R','','','','','','',302,0,10,1,0,0,0,'',14.95,0,0.00,1,226,3,0,'',0,0,0.000,'','-',0,'','-',0,'26448.png','','',0,0.000,0),(11938,26466,'QUESO EDAM BARRA','QUESO EDAM BARRA','','','','','','',302,0,11,1,0,0,0,'',8.75,0,0.00,1,274,3,0,'',0,0,0.000,'','-',0,'','-',0,'26466.png','','',0,0.000,0),(11939,26492,'JAMON CURADO HEMBRA PIEZA','JAMON CURADO HEMBRA PIEZA','','','','','','',301,0,22,1,0,0,0,'',49.00,0,0.00,0,805,3,0,'',0,0,0.000,'','-',0,'','-',0,'26492.png','','',0,0.000,0),(11940,26724,'POLVORON ALMENDRA SIN AZU','POLVORON ALMENDRA SIN AZU','','','','','','',603,0,12,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'26724.png','','',0,0.000,0),(11941,26735,'QUESO GRUYERE','QUESO GRUYERE','','','','','','',302,0,13,1,0,0,0,'',21.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'26735.png','','',0,0.000,0),(11942,26742,'QUESO ROQUEFORT','QUESO ROQUEFORT','','','','','','',302,0,14,1,0,0,0,'',15.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'26742.png','','',0,0.000,0),(11943,26744,'GAMBON 10-20','GAMBON 10-20','','','','','','',702,0,13,1,0,0,0,'',14.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'26744.png','','',0,0.000,0),(11944,26768,'PIMIENTO TRICOLOR  BANDEJ','PIMIENTO TRICOLOR  BANDEJ','','','','','','',502,0,46,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'26768.png','','',0,0.000,0),(11945,26777,'PUNTA JAMON','PUNTA JAMON','','','','','','',301,0,21,1,0,0,0,'',10.90,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'26777.png','','',0,0.000,0),(11946,26778,'HAMBURGUESA DE AVE','HAMBURGUESA DE AVE','','','','','','',208,0,8,1,0,0,0,'',2.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'26778.png','','',0,0.000,0),(11947,26782,'HAMBURGUESA MIXTA BANDEJA','HAMBURGUESA MIXTA BANDEJA','','','','','','',202,0,17,1,0,0,0,'',3.15,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'26782.png','','',0,0.000,0),(11948,26795,'QUESO GOUDA VIEJO','QUESO GOUDA VIEJO','','','','','','',302,0,15,1,0,0,0,'',15.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'26795.png','','',0,0.000,0),(11949,27188,'QUESO GOUDA SEMI','QUESO GOUDA SEMI','','','','','','',302,0,16,1,0,0,0,'',9.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'27188.png','','',0,0.000,0),(11950,27217,'JAMON SERRANO 4 ESTACIO.','JAMON SERRANO 4 ESTACIO.','','','','','','',301,0,22,1,0,0,0,'',59.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'27217.png','','',0,0.000,0),(11951,27218,'AJOS MORADOS MALLA','AJOS MORADOS MALLA','','','','','','',502,0,47,0,0,0,0,'',1.39,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'27218.png','','',0,0.000,0),(11952,27307,'TELLINA MALLA 1/2 KG','TELLINA MALLA 1/2 KG','','','','','','',702,0,14,1,0,0,0,'',9.90,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'27307.png','','',0,0.000,0),(11953,27340,'PICADA MIXTA BURGER MEAT','PICADA MIXTA BURGER MEAT','','','','','','',208,0,9,1,0,0,0,'',3.10,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'27340.png','','',0,0.000,0),(11954,27786,'EMPANADILLA AMARILLA','EMPANADILLA AMARILLA','','Ingredientes:HARINA DE TRIGO,agua,manteca,emulgentes,E-471,antioxidante,E-320,E-321,aceite de girasol,vino blanco,colorante,E-104*,maltodextrina de maiz,sal;Relleno:tomate,aceite de girasol,azucar,ATUN,HUEVO,*puede tener efectos negativos sobre la actividad y atencion de los niños;Puede contener trazas de:CRUSTÁCEOS Y DERIVADOS,SOJA,LECHE,APIO,','','','','',602,0,13,0,0,0,0,'',0.40,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'27786.png','','',0,0.000,0),(11955,27787,'EMPANADILLA MASA ROJA','EMPANADILLA MASA ROJA','','Ingredientes:HARINA DE TRIGO,agua,manteca de cerdo,emulgentes,E-471,antioxidante,E-320,E-321,aceite de girasol,vino blanco,pimenton,sal;Relleno:tomate,aceite de girasol,azucar,sal,ATUN,huevo,sal;Puede contener trazas de:CRUSTÁCEOS Y DERIVADOS,SOJA,LECHE Y DERIVADOS,APIO','','','','',602,0,14,0,0,0,0,'',0.56,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'27787.png','','',0,0.000,0),(11956,27788,'SUPER PIZZA YORK','SUPER PIZZA YORK','','Ingredientes:HARINA DE TRIGO,agua,aceite vegetal,margarina,(,grasa vegetal parcialmente hidrogenada,sal,emulgentes,E-471,E-322,aromas,acidulante,E-330,conservador,E-200,antioxidante,E-304,E-306,colorante,E-160a,),levadura,vino blanco,azucar;Relleno:tomate,aceite de girasol,carne de cerdo,fecula de patata,PROTEINAS LACTEAS,PROTEINA DE SOJA,dextrosa,LACTOSA,aromas,estabilizador,E-451,E-407,potenciador del sabor,E-621,conservador,E-250,E-252,antioxidante,E-316,colorante natural,E-120;Queso:queso mozzarella (leche);Sales fundentes:E-331,sal,ajo,oregano;Puede contener trazas de:CRUSTÁCEOS Y DERIVADOS,PESCADO,APIO','','','','',602,0,15,0,0,0,0,'',1.15,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'27788.png','','',0,0.000,0),(11957,27789,'PASTEL CARNE','PASTEL CARNE','','Ingredientes:Masa:HARINA DE TRIGO,agua,manteca de cerdo,antioxidante,E-310,sal;Relleno:carne de vacuno,aceite de girasol,HUEVO;Chorizo:carne de cerdo,azucares,fecula de patata,emulgentes,E-450,E-451,especias,antioxidante,E-301,antioxidante,E-316,E-331,conservador,E-252,colorante,E-120,proteina vegetal,ajo,pimienta;Puede contener trazas de:CRUSTÁCEOS Y DERIVADOS,PESCADO,SOJA,LECHE,APIO','','','','',602,0,16,0,0,0,0,'',1.30,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'27789.png','','',0,0.000,0),(11958,28571,'QUESO CRANDELA CON ARANDA','QUESO CRANDELA CON ARANDA','','','','','','',302,0,17,1,0,0,0,'',19.00,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'28571.png','','',0,0.000,0),(11959,28573,'QUESO CON PIÑA','QUESO CON PIÑA','','','','','','',302,0,18,1,0,0,0,'',14.95,0,0.00,1,0,3,0,'',10,0,0.000,'','-',0,'','-',0,'28573.png','','',0,0.000,0),(11960,28820,'MANTECADO CANELA','MANTECADO CANELA','','','','','','',603,0,13,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28820.png','','',0,0.000,0),(11961,28821,'ALFAJORES DE ALMENDRA','ALFAJORES DE ALMENDRA','','','','','','',603,0,14,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28821.png','','',0,0.000,0),(11962,28824,'BOLAS DE COCO','BOLAS DE COCO','','','','','','',603,0,15,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28824.png','','',0,0.000,0),(11963,28825,'MANTECADO DE COCO','MANTECADO DE COCO','','','','','','',603,0,16,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28825.png','','',0,0.000,0),(11964,28827,'HOJALDRADAS','HOJALDRADAS','','','','','','',603,0,17,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28827.png','','',0,0.000,0),(11965,28828,'MANTECADO DE LIMON','MANTECADO DE LIMON','','','','','','',603,0,18,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28828.png','','',0,0.000,0),(11966,28830,'MANTECADO DE CHOCOLATE','MANTECADO DE CHOCOLATE','','','','','','',603,0,19,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28830.png','','',0,0.000,0),(11967,28831,'POLVORONES DE ALMENDRA','POLVORONES DE ALMENDRA','','','','','','',603,0,20,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28831.png','','',0,0.000,0),(11968,28832,'CHOCOLATITO','CHOCOLATITO','','','','','','',603,0,21,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28832.png','','',0,0.000,0),(11969,28834,'PRALINE','PRALINE','','','','','','',603,0,22,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'28834.png','','',0,0.000,0),(11970,29323,'QUESO ZAMORANO D.O.','QUESO ZAMORANO D.O.','','','','','','',302,0,2,1,0,0,0,'',17.50,0,0.00,1,30,3,0,'',0,0,0.000,'','-',0,'','-',0,'29323.png','','',0,0.000,0),(11971,29391,'CALLOS DE VACUNO BANDEJA','CALLOS DE VACUNO BANDEJA','','','','','','',201,0,11,1,0,0,0,'',2.99,0,0.00,1,859,2,0,'',0,0,0.000,'','-',0,'','-',0,'29391.png','','',0,0.000,0),(11972,29405,'JAMON COCIDO EXTRA S/FOSF','JAMON COCIDO EXTRA S/FOSF','','','','','','',301,0,23,1,0,0,0,'',10.95,0,0.00,1,856,3,0,'',15,0,0.000,'','-',0,'','-',0,'29405.png','','',0,0.000,0),(11973,29431,'POLLO RELLENO AL HORNO','POLLO RELLENO AL HORNO','','','','','','',301,0,10,1,0,0,0,'',10.55,0,0.00,1,126,3,0,'',30,0,0.000,'','-',0,'','-',0,'29431.png','','',0,0.000,0),(11974,29455,'QUESO HAVARTI DOFINO','QUESO HAVARTI DOFINO','','','','','','',302,0,12,1,0,0,0,'',11.95,0,0.00,1,294,3,0,'',0,0,0.000,'','-',0,'','-',0,'29455.png','','',0,0.000,0),(11975,29961,'SOBRASADA MALLORQUINA','SOBRASADA MALLORQUINA','','','','','','',301,0,11,1,0,0,0,'',9.76,0,0.00,1,214,3,0,'',0,0,0.000,'','-',0,'','-',0,'29961.png','','',0,0.000,0),(11976,29969,'PERA PACKHAM','PERA PACKHAM','','Origen:España;Categoria:i;Variedad:packhams;Calibre:60+','','','','',501,0,25,1,0,0,0,'',1.39,0,0.00,1,81,4,0,'',0,0,0.000,'','-',0,'','-',0,'29969.png','','',0,0.000,0),(11977,30096,'PANCETA FILETES BANDEJA','PANCETA FILETES BANDEJA','','','','','','',202,0,18,1,0,0,0,'',6.50,0,0.00,1,784,2,0,'',0,0,0.000,'','-',0,'','-',0,'30096.png','','',0,0.000,0),(11978,30287,'ENSALADA 4 ESTACIONES','ENSALADA 4 ESTACIONES','','','','','','',502,0,48,0,0,0,0,'',0.65,0,0.00,1,996,4,0,'',0,0,0.000,'','-',0,'','-',0,'30287.png','','',0,0.000,0),(11979,30291,'ACELGA BOLSA','ACELGA BOLSA','','','','','','',502,0,49,0,0,0,0,'',1.00,0,0.00,1,992,4,0,'',0,0,0.000,'','-',0,'','-',0,'30291.png','','',0,0.000,0),(11980,30292,'ESPINACA BOLSA','ESPINACA BOLSA','','','','','','',502,0,50,0,0,0,0,'',1.00,0,0.00,1,991,4,0,'',0,0,0.000,'','-',0,'','-',0,'30292.png','','',0,0.000,0),(11981,30293,'ENSALADA GOURMET','ENSALADA GOURMET','','','','','','',502,0,51,0,0,0,0,'',1.49,0,0.00,1,990,4,0,'',0,0,0.000,'','-',0,'','-',0,'30293.png','','',0,0.000,0),(11982,30294,'ENSALADA MEZCLUM','ENSALADA MEZCLUM','','','','','','',502,0,52,0,0,0,0,'',1.39,0,0.00,1,988,4,0,'',0,0,0.000,'','-',0,'','-',0,'30294.png','','',0,0.000,0),(11983,30335,'HOJALDRADAS BAÑADAS CHOCO','HOJALDRADAS BAÑADAS CHOCO','','','','','','',603,0,23,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'30335.png','','',0,0.000,0),(11984,30662,'FILETE DE LOMO BANDEJA','FILETE DE LOMO BANDEJA','','','','','','',202,0,19,1,0,0,0,'',7.99,0,0.00,1,790,2,0,'',0,0,0.000,'','-',0,'','-',0,'30662.png','','',0,0.000,0),(11985,30665,'CHULETA DE LOMO BANDEJA','CHULETA DE LOMO BANDEJA','','','','','','',202,0,20,1,0,0,0,'',5.99,0,0.00,1,793,2,0,'',0,0,0.000,'','-',0,'','-',0,'30665.png','','',0,0.000,0),(11986,30712,'POLVORON TRADICIONAL S/GL','POLVORON TRADICIONAL S/GL','','','','','','',603,0,24,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'30712.png','','',0,0.000,0),(11987,30760,'QUESO AZUL CASTELLO','QUESO AZUL CASTELLO','','','','','','',302,0,15,1,0,0,0,'',13.55,0,0.00,1,955,3,0,'',0,0,0.000,'','-',0,'','-',0,'30760.png','','',0,0.000,0),(11988,30863,'SOLOMILLO DE POLLO BANDEJ','SOLOMILLO DE POLLO BANDEJ','','','','','','',205,0,19,1,0,0,0,'',8.90,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'30863.png','','',0,0.000,0),(11989,31200,'HUESO Y PIELES JAMON PATA','HUESO Y PIELES JAMON PATA','','','','','','',301,0,26,1,0,0,0,'',0.50,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'31200.png','','',0,0.000,0),(11990,31237,'PICADA VACUNO BURGER MEAT','PICADA VACUNO BURGER MEAT','','','','','','',201,0,12,1,0,0,0,'',3.10,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'31237.png','','',0,0.000,0),(11991,31241,'HUESOS COCIDO DE CERDO','HUESOS COCIDO DE CERDO','','','','','','',202,0,21,1,0,0,0,'',2.99,0,0.00,1,326,2,0,'',0,0,0.000,'','-',0,'','-',0,'31241.png','','',0,0.000,0),(11992,31269,'CEBOLLA EXTRA TUBO 1KG','CEBOLLA EXTRA TUBO 1KG','','','','','','',502,0,53,0,0,0,0,'',1.25,0,0.00,1,112,4,0,'',0,0,0.000,'','-',0,'','-',0,'31269.png','','',0,0.000,0),(11993,31692,'COSTILLA CERDO ASADA BARB','COSTILLA CERDO ASADA BARB','','','','','','',208,0,10,1,0,0,0,'',6.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'31692.png','','',0,0.000,0),(11994,31693,'COSTILLA DE CERDO ASADA M','COSTILLA DE CERDO ASADA M','','','','','','',208,0,11,1,0,0,0,'',6.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'31693.png','','',0,0.000,0),(11995,31696,'QUESO GOUDA BARRA','QUESO GOUDA BARRA','','','','','','',302,0,22,1,0,0,0,'',8.75,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'31696.png','','',0,0.000,0),(11996,31749,'TRUFAS CONGELADAS','TRUFAS CONGELADAS','','','','','','',603,0,25,0,0,0,0,'',2.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'31749.png','','',0,0.000,0),(11997,31814,'CHORIZO EXTRA VELA MINESC','CHORIZO EXTRA VELA MINESC','','','','','','',301,0,27,1,0,0,0,'',8.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'31814.png','','',0,0.000,0),(11998,31815,'JAMON COCIDO BIENSTAR BG','JAMON COCIDO BIENSTAR BG','','','','','','',301,0,28,1,0,0,0,'',8.90,0,0.00,1,0,3,0,'',15,0,0.000,'','-',0,'','-',0,'31815.png','','',0,0.000,0),(11999,31820,'BARRA GRAN RESERVA MASA M','BARRA GRAN RESERVA MASA M','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura,MASA MADRE DE TRIGO,harina de trigo fermentada;Emulgentes:E-472e;Puede contener trazas de:SOJA,LECHE Y DERIVADOS,SESAMO','','','','',601,0,15,0,0,0,0,'',0.90,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'31820.png','','',0,0.000,0),(12000,31821,'BOCATA GRAN RESERVA MASA','BOCATA GRAN RESERVA MASA','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura,MASA MADRE DE TRIGO;Emulgentes:E-472(e);Puede contener trazas de:SOJA Y DERIVADOS,LECHE Y DERIVADOS,SESAMO','','','','',601,0,16,0,0,0,0,'',0.50,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'31821.png','','',0,0.000,0),(12001,31823,'BARRA MEDITERRANEA G. RES','BARRA MEDITERRANEA G. RES','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura,MASA MADRE DE TRIGO;Emulgentes:E-472e;Contiene trazas de:SOJA Y DERIVADOS,LECHE Y DERIVADOS,SESAMO,FRUTOS DE CÁSCARA','','','','',601,0,17,0,0,0,0,'',0.80,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'31823.png','','',0,0.000,0),(12002,31825,'PAN CENTENO GRAN RESERVA','PAN CENTENO GRAN RESERVA','','Ingredientes:HARINA DE TRIGO,agua,HARINA DE CENTENO,HARINA DE TRIGO,EXTRACTO DE Malta DE CEBADA;Acidulantes:E-270,sal,levadura;Emulgentes:E-472e;Puede contener trazas de:SOJA Y DERIVADOS,LECHE Y DERIVADOS,SESAMO','','','','',601,0,18,0,0,0,0,'',1.50,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'31825.png','','',0,0.000,0),(12003,31934,'AJOS TIERNOS','AJOS TIERNOS','','','','','','',502,0,54,0,0,0,0,'',1.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'31934.png','','',0,0.000,0),(12004,32018,'COMBINADO DE PESCADO','COMBINADO DE PESCADO','','','','','','',701,0,10,1,0,0,0,'',8.80,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'32018.png','','',0,0.000,0),(12005,32050,'HAMBURGUESAS VACUNO BANDE','HAMBURGUESAS VACUNO BANDE','','','','','','',201,0,13,1,0,0,0,'',2.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'32050.png','','',0,0.000,0),(12006,32111,'DOT CHOCOLATE P-4','DOT CHOCOLATE P-4','','Contiene:GLUTEN,SOJA Y DERIVADOS,LECHE Y DERIVADOS,HUEVO Y DERIVADOS;Puede contener trazas de:FRUTOS DE CÁSCARA','','','','',602,0,17,0,0,0,0,'',1.30,0,0.00,0,0,6,0,'',3,0,0.000,'','-',0,'','-',0,'32111.png','','',0,0.000,0),(12007,32112,'DOTS AZUCAR P-4','DOTS AZUCAR P-4','','Contiene:GLUTEN,SOJA Y DERIVADOS,LECHE Y DERIVADOS,HUEVO Y DERIVADOS;Puede contener trazas de:FRUTOS DE CÁSCARA','','','','',602,0,18,0,0,0,0,'',1.30,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'32112.png','','',0,0.000,0),(12008,32327,'QUESO MEZCLA SEMICURADO','QUESO MEZCLA SEMICURADO','','','','','','',302,0,23,1,0,0,0,'',9.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'32327.png','','',0,0.000,0),(12009,32328,'QUESO MEZCLA CURADO','QUESO MEZCLA CURADO','','Ingredientes:LECHE DE VACA,LECHE DE OVEJA,LECHE DE CABRA,sal,fermentos lacticos,cuajo,cloruro calcico,lisozima;Contiene:HUEVO;Condiciones de conservacion:conservar en sitio fresco y seco','','','','',302,0,24,1,0,0,0,'',9.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'32328.png','','',0,0.000,0),(12010,32420,'SEBO Y GRASA','SEBO Y GRASA','','','','','','',208,0,12,1,0,0,0,'',0.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'32420.png','','',0,0.000,0),(12011,32457,'KIWI GOLD','KIWI GOLD','','Variedad:sungold;Origen:Italia;Categoria:i;Calibre:22/25','','','','',501,0,3,1,0,0,0,'',3.49,0,0.00,1,5,4,0,'',0,0,0.000,'','-',0,'','-',0,'32457.png','','',0,0.000,0),(12012,32458,'LLONGANISSA DE PAGES','LLONGANISSA DE PAGES','','Ingredientes:magro de cerdo,panceta de cerdo ibérico,sal,LACTOSA,PROTEINAS LACTEAS,especias,azucar,dextrina,vino;Antioxidantes:E-301,E-331;Conservadores:E-250,E-252;Colorantes:E-120;Condiciones de conservacion:conservar en sitio fresc y seco','','','','',301,0,29,1,0,0,0,'',14.90,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'32458.png','','',0,0.000,0),(12013,32499,'JAMON COCIDO BRASEADO','JAMON COCIDO BRASEADO','','','','','','',301,0,30,1,0,0,0,'',7.99,0,0.00,1,0,3,0,'',15,0,0.000,'','-',0,'','-',0,'32499.png','','',0,0.000,0),(12014,32696,'PERA CONFERENCIA CESTA','PERA CONFERENCIA CESTA','','','','','','',501,0,32,0,0,0,0,'',1.39,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'32696.png','','',0,0.000,0),(12015,33193,'QUESO SERVILLETA CABRA MA','QUESO SERVILLETA CABRA MA','','','','','','',302,0,25,1,0,0,0,'',15.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'33193.png','','',0,0.000,0),(12016,33252,'MEJILLON','MEJILLON','','','','','','',702,0,15,1,0,0,0,'',2.60,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'33252.png','','',0,0.000,0),(12017,33255,'QUESO TOLKO PIMIENTA','QUESO TOLKO PIMIENTA','','','','','','',302,0,26,1,0,0,0,'',14.95,0,0.00,1,0,3,0,'',10,0,0.000,'','-',0,'','-',0,'33255.png','','',0,0.000,0),(12018,33309,'PAN BARRA ACEITE','PAN BARRA ACEITE','','Ingredientes:HARINA DE TRIGO,agua,aceite de oliva,sal,levadura;Emulgentes:E-472(e);Puede contener trazas de:SOJA Y DERIVADOS,SESAMO,LECHE Y DERIVADOS,FRUTOS DE CÁSCARA','','','','',601,0,19,0,0,0,0,'',0.50,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'33309.png','','',0,0.000,0),(12019,33330,'QUESO FRESCO PEQ.SERVIL.','QUESO FRESCO PEQ.SERVIL.','','','','','','',302,0,27,1,0,0,0,'',8.75,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'33330.png','','',0,0.000,0),(12020,33416,'BACON','BACON','','','','','','',301,0,31,1,0,0,0,'',9.95,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'33416.png','','',0,0.000,0),(12021,33615,'MORCILLA DE BURGOS BANDEJ','MORCILLA DE BURGOS BANDEJ','','','','','','',210,0,11,1,0,0,0,'',1.59,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'33615.png','','',0,0.000,0),(12022,33690,'GAMBA SALADA COCIDA','GAMBA SALADA COCIDA','','','','','','',702,0,16,1,0,0,0,'',14.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'33690.png','','',0,0.000,0),(12023,33937,'ESCALOPE DE POLLO BANDEJA','ESCALOPE DE POLLO BANDEJA','','','','','','',205,0,20,1,0,0,0,'',4.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'33937.png','','',0,0.000,0),(12024,34316,'MELON CORTADO EXTRA','MELON CORTADO EXTRA','','Origen:España;Variedad:piel de sapo;Categoria:i;Calibre:6','','','','',501,0,33,0,0,0,0,'',2.15,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'34316.png','','',0,0.000,0),(12025,34331,'ZANAHORIA BOLSA 1/2KG','ZANAHORIA BOLSA 1/2KG','','','','','','',502,0,55,0,0,0,0,'',0.69,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'34331.png','','',0,0.000,0),(12026,34446,'ESCALOPIN DE VACUNO','ESCALOPIN DE VACUNO','','','','','','',201,0,14,1,0,0,0,'',10.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'34446.png','','',0,0.000,0),(12027,34456,'QUESO DE OVEJA CURADO','QUESO DE OVEJA CURADO','','Ingredientes:LECHE DE OVEJA,sal,fermentos lacticos,cuajo,cloruro calcico,lisozima;Contiene:HUEVO Y DERIVADOS;Condiciones de conservacion:conservar en sitio fresco y seco','','','','',302,0,28,1,0,0,0,'',12.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'34456.png','','',0,0.000,0),(12028,34517,'UVA BLANCA SIN SEMILLA TA','UVA BLANCA SIN SEMILLA TA','','','','','','',501,0,34,0,0,0,0,'',2.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'34517.png','','',0,0.000,0),(12029,34548,'LECHUGA HOJA DE ROBLE','LECHUGA HOJA DE ROBLE','','Origen:España','','','','',502,0,56,0,0,0,0,'',1.29,0,0.00,1,410,4,0,'',0,0,0.000,'','-',0,'','-',0,'34548.png','','',0,0.000,0),(12030,34583,'LONGANIZA POLLO SIN TRIPA','LONGANIZA POLLO SIN TRIPA','','','','','','',210,0,12,1,0,0,0,'',2.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'34583.png','','',0,0.000,0),(12031,34587,'QUESO MANCHEGO  D.O. CURA','QUESO MANCHEGO  D.O. CURA','','Ingredientes:LECHE DE OVEJA,fermentos lacticos,sal,cuajo,lisozima;Contiene:LECHE Y DERIVADOS,HUEVO Y DERIVADOS;Condiciones de conservacion:conservar en fresco y seco','','','','',302,0,29,1,0,0,0,'',16.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'34587.png','','',0,0.000,0),(12032,34611,'HORTALISTAS CLASICAS MICR','HORTALISTAS CLASICAS MICR','','','','','','',502,0,57,0,0,0,0,'',1.60,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'34611.png','','',0,0.000,0),(12033,34639,'MORAS','MORAS','','','','','','',501,0,35,0,0,0,0,'',2.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'34639.png','','',0,0.000,0),(12034,34642,'QUESO CHEDDAR MATURE BARR','QUESO CHEDDAR MATURE BARR','','','','','','',302,0,30,1,0,0,0,'',13.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'34642.png','','',0,0.000,0),(12035,34670,'JUDIA FINA BANDEJA','JUDIA FINA BANDEJA','','','','','','',502,0,58,0,0,0,0,'',1.39,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'34670.png','','',0,0.000,0),(12036,34672,'QUESO MEZCLA BARRA TIERNO','QUESO MEZCLA BARRA TIERNO','','','','','','',302,0,31,1,0,0,0,'',8.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'34672.png','','',0,0.000,0),(12037,34681,'MINI SALADAS SURTIDO','MINI SALADAS SURTIDO','','Ingredientes:HARINA DE TRIGO,margarina,agua,azucar,levadura,sal,ATUN,aceite de girasol,tomate,PAN RALLADO,almidon de maiz;Almidón modificado:E-1414,SUERO LACTEO,LECHE,grasa vegetal,carne de cerdo,almidon de patata,dextrosa,PROTEINA DE SOJA,aromas,especias,aceite de girasol,fecula de patata,CLARA DE HUEVO,caseinato de sodio,linaza,SESAMO,SALVADO DE AVENA,semilla de amapola,jarabe de glucosa,HUEVO ENTERO EN POLVO;Emulgentes:E-471,E-339;Acidulantes:E-330;Estabilizantes:E-451,E-401,E-461,E-407,E-410,E-412,E-415,E-450;Antioxidante:E-301;Conservadores:E-250;Colorantes:E-120,E-171;Potenciador del sabor:E-621;Puede contener trazas de:FRUTOS DE CÁSCARA','','','','',602,0,19,0,0,0,0,'',7.68,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'34681.png','','',0,0.000,0),(12038,34682,'MINI SALADOS SURTIDO','MINI SALADOS SURTIDO','','Ingredientes:HARINA DE TRIGO,margarina,agua,azucar,levadura,sal,ATUN,aceite de girasol,tomate,PAN RALLADO,almidon de maiz;Almidón modificado:E-1414,SUERO LACTEO,LECHE,grasa vegetal,carne de cerdo,almidon de patata,dextrosa,PROTEINA DE SOJA,aromas,especias,aceite de girasol,fecula de patata,CLARA DE HUEVO,caseinato de sodio,linaza,SESAMO,SALVADO DE AVENA,semilla de amapola,jarabe de glucosa,HUEVO ENTERO EN POLVO;Emulgentes:E-471,E-339;Acidulantes:E-330;Estabilizadores:E-451,E-401,E-461,E-407,E-410,E-412,E-415,E-450;Antioxidante:E-301;Conservadores:E-250;Colorantes:E-120,E-171;Potenciador del sabor:E-621;Puede contener trazas de:FRUTOS DE CÁSCARA','','','','',602,0,20,0,0,0,0,'',1.19,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'34682.png','','',0,0.000,0),(12039,34836,'NUECES','NUECES','','Origen:EEUU;Categoria:i;Calibre:+32;Variedad:yumbo','','','','',505,0,5,1,0,0,0,'',4.49,0,0.00,1,64,4,0,'',0,0,0.000,'','-',0,'','-',0,'34836.png','','',0,0.000,0),(12040,34918,'CHAMPIÑON PIE CORTADO','CHAMPIÑON PIE CORTADO','','','','','','',502,0,59,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'34918.png','','',0,0.000,0),(12041,34919,'SETAS','SETAS','','','','','','',502,0,60,0,0,0,0,'',2.05,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'34919.png','','',0,0.000,0),(12042,34975,'BOLITA COCO ARTESANO','BOLITA COCO ARTESANO','','','','','','',603,0,26,0,0,0,0,'',1.65,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'34975.png','','',0,0.000,0),(12043,35111,'MOUSSE CANARD SETAS Y CHA','MOUSSE CANARD SETAS Y CHA','','','','','','',301,0,32,1,0,0,0,'',21.75,0,0.00,1,0,3,0,'',10,0,0.000,'','-',0,'','-',0,'35111.png','','',0,0.000,0),(12044,35112,'QUESO CURADO RESERVA 12 M','QUESO CURADO RESERVA 12 M','','','','','','',302,0,32,1,0,0,0,'',15.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'35112.png','','',0,0.000,0),(12045,35199,'CEBOLLAS','CEBOLLAS','','Origen:España;Variedad:grano;Calibre:70/90;Categoria:1º','','','','',502,0,12,1,0,0,0,'',1.15,0,0.00,1,52,4,0,'',0,0,0.000,'','-',0,'','-',0,'35199.png','','',0,0.000,0),(12046,35260,'KIWI VALENCIANO','KIWI VALENCIANO','','Origen:valencia;Categoria:primera;Variedad:hayward;Calibre:18/20','','','','',501,0,21,1,0,0,0,'',2.49,0,0.00,1,70,4,0,'',0,0,0.000,'','-',0,'','-',0,'35260.png','','',0,0.000,0),(12047,35273,'PINCHOS DE CERDO BANDEJA','PINCHOS DE CERDO BANDEJA','','','','','','',202,0,22,1,0,0,0,'',3.30,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'35273.png','','',0,0.000,0),(12048,35274,'CARPACCIO DE VACUNO','CARPACCIO DE VACUNO','','','','','','',201,0,15,1,0,0,0,'',3.25,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'35274.png','','',0,0.000,0),(12049,35328,'JENGIBRE','JENGIBRE','','','','','','',502,0,62,0,0,0,0,'',1.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35328.png','','',0,0.000,0),(12050,35340,'QUESO FRESCO CINCHO CON S','QUESO FRESCO CINCHO CON S','','','','','','',302,0,33,1,0,0,0,'',8.75,0,0.00,1,0,3,0,'',2,0,0.000,'','-',0,'','-',0,'35340.png','','',0,0.000,0),(12051,35378,'FRIKORN MULTICEREAL','FRIKORN MULTICEREAL','','Ingredientes:HARINA DE TRIGO,agua,HARINA DE TRIGO,HARINA DE CENTENO,HARINA DE CEBADA,pipas de girasol,linaza,SOJA,GLUTEN DE TRIGO,SALVADO DE TRIGO,maiz,azucar,sal,aceite de palma;Emulgentes:E-472e,mijo,levadura,sal,citrato potasico,GLUTEN,vinagre,CEBADA,HARINA INTEG.CENTENO;Acidulantes:ácido láctico;Emulgentes:E-472e;Puede contener trazas de:HUEVO,SOJA,LECHE,SESAMO,ALTRAMUCES','','','','',601,0,20,0,0,0,0,'',1.65,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'35378.png','','',0,0.000,0),(12052,35380,'FRIKORN RUSTIKAL','FRIKORN RUSTIKAL','','Ingredientes:HARINA DE TRIGO,agua,HARINA DE CENTENO,linaza,SESAMO,sal,COPOS DE AVENA,pipas de girasol,SALVADO DE TRIGO,HARINA DE Malta DE TRIGO,cascaras de guisantes,semola de maiz,SUERO DE LECHE EN POLVO;Acidulantes:E-270;Emulgentes:E-472e,pipas de calabaza,levadura,sal,citrato potasico,vinagre;Acidulantes:acido lactico;Puede contener trazas de:HUEVO,SOJA,ALTRAMUCES,SESAMO','','','','',601,0,21,0,0,0,0,'',1.65,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'35380.png','','',0,0.000,0),(12053,35394,'LAZOS AZUCAR','LAZOS AZUCAR','','','','','','',602,0,21,0,0,0,0,'',1.35,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'35394.png','','',0,0.000,0),(12054,35396,'PALMERAS AZUCAR','PALMERAS AZUCAR','','','','','','',602,0,22,0,0,0,0,'',1.35,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'35396.png','','',0,0.000,0),(12055,35588,'PAN BARRA 2U','PAN BARRA 2U','','Ingredientes:HARINA DE TRIGO,agua,levadura,sal;Emulgentes:E-472(e);Puede contener trazas de:SOJA,LECHE,SESAMO','','','','',601,0,22,0,0,0,0,'',0.84,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'35588.png','','',0,0.000,0),(12056,35603,'COL LISA REPOLLO KILO','COL LISA REPOLLO KILO','','Origen:España;Variedad:lisa;Calibre:8;Categoria:1º','','','','',502,0,63,0,0,0,0,'',1.19,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35603.png','','',0,0.000,0),(12057,35623,'RUCULA BOLSA','RUCULA BOLSA','','','','','','',502,0,64,0,0,0,0,'',1.35,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35623.png','','',0,0.000,0),(12058,35624,'ENSALADA SUPREME BOLSA','ENSALADA SUPREME BOLSA','','','','','','',502,0,65,0,0,0,0,'',1.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35624.png','','',0,0.000,0),(12059,35625,'BROTES DE SOJA','BROTES DE SOJA','','','','','','',502,0,66,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35625.png','','',0,0.000,0),(12060,35693,'QUESO FRESCO DE CABRA','QUESO FRESCO DE CABRA','','','','','','',302,0,34,1,0,0,0,'',11.55,0,0.00,1,0,3,0,'',2,0,0.000,'','-',0,'','-',0,'35693.png','','',0,0.000,0),(12061,35741,'LOMO CINTA PIEZA ENTERA/M','LOMO CINTA PIEZA ENTERA/M','','','','','','',202,0,23,1,0,0,0,'',6.80,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'35741.png','','',0,0.000,0),(12062,35876,'HAMBURGUESA VACUNO 1/4 LI','HAMBURGUESA VACUNO 1/4 LI','','','','','','',201,0,16,1,0,0,0,'',2.35,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'35876.png','','',0,0.000,0),(12063,35938,'UVA NEGRA SIN SEMILLAS TA','UVA NEGRA SIN SEMILLAS TA','',',','','','','',501,0,37,0,0,0,0,'',2.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35938.png','','',0,0.000,0),(12064,35957,'ENSALADA LUXURE','ENSALADA LUXURE','','','','','','',502,0,67,0,0,0,0,'',1.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35957.png','','',0,0.000,0),(12065,35958,'ENSALADA DELICIOUS','ENSALADA DELICIOUS','','','','','','',502,0,68,0,0,0,0,'',1.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35958.png','','',0,0.000,0),(12066,35968,'HIGADO DE CORDERO','HIGADO DE CORDERO','','','','','','',203,0,11,1,0,0,0,'',8.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'35968.png','','',0,0.000,0),(12067,35977,'RABANITOS','RABANITOS','','','','','','',502,0,69,0,0,0,0,'',1.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'35977.png','','',0,0.000,0),(12068,35981,'PAN PULGA','PAN PULGA','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura;Emulgentes:E-472e;Puede contener trazas de:SOJA Y DERIVADOS,LECHE Y DERIVADOS,SESAMO,FRUTOS DE CÁSCARA','','','','',601,0,23,0,0,0,0,'',0.17,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'35981.png','','',0,0.000,0),(12069,36001,'ARREGLO COCIDO','ARREGLO COCIDO','','','','','','',502,0,70,0,0,0,0,'',1.39,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36001.png','','',0,0.000,0),(12070,36016,'PATATAS MICROONDAS','PATATAS MICROONDAS','','','','','','',502,0,71,0,0,0,0,'',0.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36016.png','','',0,0.000,0),(12071,36066,'CARRILLADA DE CERDO GRANE','CARRILLADA DE CERDO GRANE','','','','','','',202,0,24,1,0,0,0,'',9.10,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'36066.png','','',0,0.000,0),(12072,36075,'COGOLLOS','COGOLLOS','','','','','','',502,0,72,0,0,0,0,'',1.15,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36075.png','','',0,0.000,0),(12073,36225,'SOLOMILLO DE PAVO','SOLOMILLO DE PAVO','','','','','','',205,0,21,1,0,0,0,'',7.75,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'36225.png','','',0,0.000,0),(12074,36243,'MANZANA GOLDEN SELECCION','MANZANA GOLDEN SELECCION','','Origen:Italia;Categoria:i;Calibre:80/85;Variedad:golden delicius','','','','',501,0,4,1,0,0,0,'',1.99,0,0.00,1,7,4,0,'',0,0,0.000,'','-',0,'','-',0,'36243.png','','',0,0.000,0),(12075,36283,'JAMON SERRANO 4 ESTAC+LOT','JAMON SERRANO 4 ESTAC+LOT','','','','','','',301,0,33,1,0,0,0,'',75.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'36283.png','','',0,0.000,0),(12076,36323,'COLES DE BRUSELAS','COLES DE BRUSELAS','','','','','','',502,0,73,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36323.png','','',0,0.000,0),(12077,36355,'SALMON AHUMADO NORUEGO PL','SALMON AHUMADO NORUEGO PL','','','','','','',301,0,34,1,0,0,0,'',29.95,0,0.00,1,0,3,0,'',9,0,0.000,'','-',0,'','-',0,'36355.png','','',0,0.000,0),(12078,36360,'SALCHICHON GREVOL','SALCHICHON GREVOL','','','','','','',301,0,35,1,0,0,0,'',8.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'36360.png','','',0,0.000,0),(12079,36405,'PAVO NACIONAL 5,5-7 KG','PAVO NACIONAL 5,5-7 KG','','','','','','',205,0,22,1,0,0,0,'',3.90,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'36405.png','','',0,0.000,0),(12080,36411,'HIERBA BUENA MACETA','HIERBA BUENA MACETA','','','','','','',502,0,74,0,0,0,0,'',1.49,0,0.00,0,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36411.png','','',0,0.000,0),(12081,36412,'PEREJIL MACETA','PEREJIL MACETA','','','','','','',502,0,75,0,0,0,0,'',1.49,0,0.00,0,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36412.png','','',0,0.000,0),(12082,36423,'FRUTOS EXOTICOS BANDEJA','FRUTOS EXOTICOS BANDEJA','','','','','','',501,0,39,0,0,0,0,'',6.75,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36423.png','','',0,0.000,0),(12083,36454,'CORAZON DE COGOLLOS','CORAZON DE COGOLLOS','','','','','','',502,0,76,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36454.png','','',0,0.000,0),(12084,36519,'MANDARINA MALLA','MANDARINA MALLA','','','','','','',501,0,40,0,0,0,0,'',2.10,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'36519.png','','',0,0.000,0),(12085,36575,'TORTA CASERA','TORTA CASERA','','Ingredientes:HARINA DE TRIGO,azucar,aceite de girasol,HUEVO,aromas,SUERO DE LECHE EN POLVO;Gasificantes:difosfato sódico,bicarbonato sódico,almidón de patata,carbonato cálcico;Estabilizadores:goma de guar,carbonato cálcico;Emulgentes:mono y digliceridos,esteres de propanodiol,sales sódicas de ácidos grasos;Humectantes:sorbitol,almidon de maiz,Ácido ortofosfórico,enzimas,carbonato cálcico,sal;Puede contener trazas de:SOJA,FRUTOS DE CÁSCARA','','','','',602,0,23,0,0,0,0,'',1.00,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'36575.png','','',0,0.000,0),(12086,36576,'TORTA CASERA LIMON','TORTA CASERA LIMON','','Ingredientes:HARINA DE TRIGO,azucar,aceite de girasol,HUEVO,aromas,SUERO DE LECHE EN POLVO;Gasificantes:difosfato disódico,bicarbonato sódico,almidon de patata,carbonato cálcico;Estabilizadores:goma de guar,carbonato de calcio;Emulgentes:mono y diglicéridos,esteres de propanodiol,sales sódicas de ácidos grasos;Humectantes:sorbitol,almidon de maiz,Ácido ortofosfórico,enzimas,carbonato cálcico,sal;Puede contener trazas de:SOJA,FRUTOS DE CÁSCARA,SESAMO','','','','',602,0,24,0,0,0,0,'',1.00,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'36576.png','','',0,0.000,0),(12087,36577,'MOJICON 4U','MOJICON 4U','','Ingredientes:HARINA DE TRIGO,azucar,aceite de girasol,HUEVO,aroma,SUERO DE LECHE EN POLVO;Gasificantes:difosfato disódico,bicarbonato sódico,almidon de patata,carbonato cálcico;Estabilizadores:goma de guar,carbonato cálcico;Emulgentes:mono y diglicéridos,esteres de propanodiol,sales sódicas de ácidos grasos;Humectantes:sorbitol,almidon de maiz,Ácido ortofosfórico,enzimas,carbonato cálcico,sal;Puede contener trazas de:SOJA,FRUTOS DE CÁSCARA','','','','',602,0,25,0,0,0,0,'',1.00,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'36577.png','','',0,0.000,0),(12088,36578,'MOJICON CHOCOLATE 4U','MOJICON CHOCOLATE 4U','','Ingredientes:HARINA DE TRIGO,aceite de girasol,azucar,HUEVO,cacao en polvo,aromas;Humectantes:glicerina,sorbitol,almidon de maiz,Ácido ortofofórico,enzimas,carbonato cálcico,SUERO DE LECHE EN POLVO;Gasificantes:difosfato disódico,bicarbonato sódico,almidon de patata,carbonato cálcio,sal;Puede contener trazas de:SOJA,FRUTOS DE CÁSCARA','','','','',602,0,26,0,0,0,0,'',1.00,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'36578.png','','',0,0.000,0),(12089,36636,'HIGADO DE CERDO','HIGADO DE CERDO','','','','','','',202,0,25,1,0,0,0,'',2.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'36636.png','','',0,0.000,0),(12090,37280,'JAMON CEBO IBERICO PIEZA','JAMON CEBO IBERICO PIEZA','','','','','','',301,0,36,1,0,0,0,'',129.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'37280.png','','',0,0.000,0),(12091,37287,'LACON EXTRA ASADO LEÑA','LACON EXTRA ASADO LEÑA','','','','','','',301,0,37,1,0,0,0,'',8.95,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'37287.png','','',0,0.000,0),(12092,37314,'ENSALADA 4 ESTACIONES MAX','ENSALADA 4 ESTACIONES MAX','','','','','','',502,0,77,0,0,0,0,'',1.10,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37314.png','','',0,0.000,0),(12093,37357,'ENSALADA MEDITERRANEA','ENSALADA MEDITERRANEA','','','','','','',502,0,78,0,0,0,0,'',1.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37357.png','','',0,0.000,0),(12094,37368,'JAMON CURADO DUROC100% PI','JAMON CURADO DUROC100% PI','','','','','','',301,0,38,1,0,0,0,'',109.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'37368.png','','',0,0.000,0),(12095,37372,'ALMEJA CHIRLA MALLA 1/2 K','ALMEJA CHIRLA MALLA 1/2 K','','','','','','',702,0,17,1,0,0,0,'',4.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'37372.png','','',0,0.000,0),(12096,37460,'ENSALADA VARIADA','ENSALADA VARIADA','','','','','','',502,0,79,0,0,0,0,'',1.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37460.png','','',0,0.000,0),(12097,37499,'HOJALDRE DE PAVO Y CAMEMB','HOJALDRE DE PAVO Y CAMEMB','','Ingredientes:Masa:HARINA DE TRIGO;Margarina:grasa de palma,ACEITE DE SOJA,aceite de girasol,agua;Emulgentes:E-471;Acidulantes:E-330,aromas;Conservantes:E-202;Colorantes:E-160a,agua,levadura,HUEVO ENTERO EN POLVO,dextrosa,sal,LECHE EN POLVO;Relleno:Crema de queso:queso fresco (leche),camembert (leche),agua,MANTEQUILLA (LECHE),queso fundido (leche),PROTEINAS LACTEAS,sal;Estabilizantes:E-452,E-450;Antioxidantes:E-341,E-339,E-331;Dados de pavo:pechuga de pavo,agua,fecula de patata,azucares;Estabilizantes:E-407,E-450,E-451;Antioxidantes:E-301,E-331;Potenciador del sabor:E-621;Conservadores:E-250,PROTEINA DE SOJA,sal,aroma,CLARA DE HUEVO,almidon de maiz;Puede contener trazas de:FRUTOS DE CÁSCARA','','','','',602,0,27,0,0,0,0,'',0.85,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'37499.png','','',0,0.000,0),(12098,37501,'CROISSANT VIENES CEREALES','CROISSANT VIENES CEREALES','','Ingredientes:HARINA DE TRIGO,agua,MANTEQUILLA (LECHE),mix cereales,HARINA INTEGRAL,semillas de lino marrón,HARINA Malta CEBADA,semillas de girasol,semilla de chía,azucar,levadura,sal,GLUTEN;Cobertura:semillas de lino dorado,semillas de lino marrón,COPOS DE AVENA,semillas de girasol,HUEVO,dextrosa;Puede contener trazas de:SOJA,FRUTOS DE CÁSCARA,SESAMO','','','','',602,0,28,0,0,0,0,'',0.75,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'37501.png','','',0,0.000,0),(12099,37568,'BOCADITO BOMBON CRUJIENTE','BOCADITO BOMBON CRUJIENTE','','','','','','',603,0,27,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'37568.png','','',0,0.000,0),(12100,37571,'CALZONE QUESO/TOMATE','CALZONE QUESO/TOMATE','','Ingredientes:Masa:HARINA DE TRIGO;Margarina:grasa de palma,agua,ACEITE DE SOJA,aceite de girasol;Emulgentes:E-471,sal;Acidulantes:E-330,aromas;Antioxidantes:E-306,E-304;Colorantes:E-160a,agua,sal,GLUTEN;Emulgentes:E-471;Relleno:QUESO RALLADO (LECHE),QUESO EN POLVO (LECHE);Pure de tomate:SUERO LACTEO;Almidón modificado:E-1414,almidon de maiz,grasa de palma,sal;Estabilizantes:E-412,E-401,oregano;Cobertura:HUEVO LíQUIDO;Puede contener trazas de:PESCADO Y DERIVADOS,SOJA,MOSTAZA,SESAMO','','','','',602,0,29,0,0,0,0,'',0.90,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'37571.png','','',0,0.000,0),(12101,37588,'QUESO GRANA PADANO','QUESO GRANA PADANO','','','','','','',302,0,35,1,0,0,0,'',21.45,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'37588.png','','',0,0.000,0),(12102,37637,'BROCOLI','BROCOLI','','','','','','',502,0,80,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37637.png','','',0,0.000,0),(12103,37682,'MANZANA ROYAL GALA BOLSA','MANZANA ROYAL GALA BOLSA','','','','','','',501,0,41,0,0,0,0,'',2.35,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37682.png','','',0,0.000,0),(12104,37732,'MUSLO DE PATO CONFITADO','MUSLO DE PATO CONFITADO','','','','','','',208,0,13,1,0,0,0,'',6.35,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'37732.png','','',0,0.000,0),(12105,37757,'TOMATE CHERRY KUMATO','TOMATE CHERRY KUMATO','','','','','','',502,0,81,0,0,0,0,'',1.39,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37757.png','','',0,0.000,0),(12106,37775,'BONIATO BLANCO','BONIATO BLANCO','','Categoria:i;Origen:España;Variedad:blanco;Calibre:300 +','','','','',502,0,6,1,0,0,0,'',2.29,0,0.00,1,23,4,0,'',0,0,0.000,'','-',0,'','-',0,'37775.png','','',0,0.000,0),(12107,37797,'ZUMO NARANJA RECIEN EXPRI','ZUMO NARANJA RECIEN EXPRI','','','','','','',501,0,42,0,0,0,0,'',1.90,0,0.00,0,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37797.png','','',0,0.000,0),(12108,37798,'ZUMO NARANJA RECIEN EXPRI','ZUMO NARANJA RECIEN EXPRI','','','','','','',501,0,43,0,0,0,0,'',2.99,0,0.00,0,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37798.png','','',0,0.000,0),(12109,37800,'HIGOS SECOS','HIGOS SECOS','','Variedad:cuello dama;Categoria:i;Origen:España;Calibre:80/100','','','','',505,0,3,1,0,0,0,'',6.49,0,0.00,1,37,4,0,'',0,0,0.000,'','-',0,'','-',0,'37800.png','','',0,0.000,0),(12110,37852,'LICHIS','LICHIS','','','','','','',501,0,44,0,0,0,0,'',2.09,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37852.png','','',0,0.000,0),(12111,37874,'BANANAS','BANANAS','','Origen:mexico;Variedad:cavendish;Calibre:19+;Categoria:i','','','','',501,0,7,1,0,0,0,'',1.39,0,0.00,1,10,4,0,'',0,0,0.000,'','-',0,'','-',0,'37874.png','','',0,0.000,0),(12112,37950,'AJOS SECOS BLANCOS','AJOS SECOS BLANCOS','','','','','','',502,0,83,0,0,0,0,'',1.29,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'37950.png','','',0,0.000,0),(12113,38000,'PANES GRANEL','PANES GRANEL','','','','','','',601,0,1,1,0,0,0,'',2.60,0,0.00,1,1,6,0,'',0,0,0.000,'','-',0,'','-',0,'38000.png','','',0,0.000,0),(12114,38004,'PAN BURGUER','PAN BURGUER','','Ingredientes:HARINA DE TRIGO,agua,aceite de girasol,GLUTEN;Emulgentes:E-472e,levadura,SóLIDOS LACTEOS,sal;Puede contener trazas de:SOJA Y DERIVADOS,SESAMO','','','','',601,0,25,0,0,0,0,'',0.25,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38004.png','','',0,0.000,0),(12115,38005,'BOLLERIA GRANEL','BOLLERIA GRANEL','','GLUTEN,SOJA,FRUTOS DE CÁSCARA,LECHE,SESAMO','','','','',602,0,1,1,0,0,0,'',7.50,0,0.00,1,5,6,0,'',0,0,0.000,'','-',0,'','-',0,'38005.png','','',0,0.000,0),(12116,38034,'BARRA RUSTICA REBANADA','BARRA RUSTICA REBANADA','','','','','','',601,0,26,0,0,0,0,'',1.00,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38034.png','','',0,0.000,0),(12117,38035,'BARRA RUSTICA MULTICEREAL','BARRA RUSTICA MULTICEREAL','','','','','','',601,0,27,0,0,0,0,'',1.19,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38035.png','','',0,0.000,0),(12118,38045,'FLAUTAS POLLO Y QUESO BAN','FLAUTAS POLLO Y QUESO BAN','','','','','','',208,0,14,1,0,0,0,'',2.05,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'38045.png','','',0,0.000,0),(12119,38046,'FLAUTAS QUESO Y BACON BAN','FLAUTAS QUESO Y BACON BAN','','','','','','',208,0,15,1,0,0,0,'',2.05,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'38046.png','','',0,0.000,0),(12120,38047,'SURTIDO EMBUTIDO BARBACOA','SURTIDO EMBUTIDO BARBACOA','','','','','','',210,0,13,1,0,0,0,'',8.35,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'38047.png','','',0,0.000,0),(12121,38048,'MORCILLA CEBOLLA OREADA B','MORCILLA CEBOLLA OREADA B','','','','','','',210,0,14,1,0,0,0,'',7.55,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'38048.png','','',0,0.000,0),(12122,38121,'(GRANEL) TRIANGULO CABELL','(GRANEL) TRIANGULO CABELL','','','','','','',603,0,28,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38121.png','','',0,0.000,0),(12123,38122,'(GRANEL) TRIANGULO CREMA','(GRANEL) TRIANGULO CREMA','','','','','','',603,0,29,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38122.png','','',0,0.000,0),(12124,38123,'(GRANEL) CAÑA CREMA ZEBRA','(GRANEL) CAÑA CREMA ZEBRA','','','','','','',603,0,30,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38123.png','','',0,0.000,0),(12125,38124,'(GRANEL) CUADRADITO CACAO','(GRANEL) CUADRADITO CACAO','','','','','','',603,0,31,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38124.png','','',0,0.000,0),(12126,38276,'PECHUGA PAVO ADOBADA BJA.','PECHUGA PAVO ADOBADA BJA.','','','','','','',205,0,23,1,0,0,0,'',9.35,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'38276.png','','',0,0.000,0),(12127,38330,'PECHUGA POLLO','PECHUGA POLLO','','','','','','',301,0,39,1,0,0,0,'',7.99,0,0.00,1,0,3,0,'',15,0,0.000,'','-',0,'','-',0,'38330.png','','',0,0.000,0),(12128,38420,'LOMO BAJO VACUNO SKIN','LOMO BAJO VACUNO SKIN','','','','','','',201,0,17,1,0,0,0,'',16.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'38420.png','','',0,0.000,0),(12129,38477,'HUESOS DE LACON','HUESOS DE LACON','','','','','','',301,0,40,1,0,0,0,'',0.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'38477.png','','',0,0.000,0),(12130,38501,'CEBOLLA DULCE','CEBOLLA DULCE','','Origen:España;Variedad:dulce;Calibre:60/80;Categoria:1º','','','','',502,0,16,1,0,0,0,'',2.40,0,0.00,1,68,4,0,'',0,0,0.000,'','-',0,'','-',0,'38501.png','','',0,0.000,0),(12131,38547,'PAN GRANO ENTERO REBANADO','PAN GRANO ENTERO REBANADO','','','','','','',601,0,28,0,0,0,0,'',1.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38547.png','','',0,0.000,0),(12132,38548,'PAN REDONDO CENTENO REBAN','PAN REDONDO CENTENO REBAN','','','','','','',601,0,29,0,0,0,0,'',1.95,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38548.png','','',0,0.000,0),(12133,38571,'CEBOLLA MORADA','CEBOLLA MORADA','','Origen:España;Variedad:roja;Calibre:60/80;Categoria:1º','','','','',502,0,17,1,0,0,0,'',2.75,0,0.00,1,69,4,0,'',0,0,0.000,'','-',0,'','-',0,'38571.png','','',0,0.000,0),(12134,38700,'BERLINA RELLENA CHOCOLATE','BERLINA RELLENA CHOCOLATE','','Ingredientes:Masa:HARINA DE TRIGO,agua,grasa de palma,azucar,levadura,aceite de nabina,dextrosa,ALMIDON DE TRIGO;Gasificantes:E-450,E-500ii,sal,CLARA DE HUEVO;Emulgentes:E-471,E-481,E-472e,SUERO LACTEO;Estabilizadores:E-412,E-466,E-415;Antioxidante:E-300,LECHE DESNATADA;Colorantes:E-160(a);Relleno:azucar,aceite de girasol,grasa de palma,grasa de palmiste,cacao desgrasado en polvo,SUERO LACTEO,AVELLANAS;Emulgentes:E-322,aromas;Cobertura:azucar,grasa de coco,grasa de palmiste,cacao desgrasado en polvo,LECHE ENTERA EN POLVO;Emulgentes:E-322,aromas naturales;Contiene:SOJA,producto descongelado','','','','',602,0,31,0,0,0,0,'',0.65,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'38700.png','','',0,0.000,0),(12135,38701,'PAN FLAUTIN CEREALES MASA','PAN FLAUTIN CEREALES MASA','','Ingredientes:HARINA DE TRIGO,agua,HARINA DE CENTENO,HARINA DE TRIGO,mijo pelado,lino marron,semillas de girasol,SESAMO,harina de arroz,GLUTEN,CEBADA,MASA MADRE DE TRIGO,sal,levadura,HARINA DE Malta DE TRIGO;Emulgentes:E-472e;Cobertura:alfordón,lino dorado,lino marron,SESAMO,semillas de girasol;Puede contener trazas de:SOJA,ALTRAMUCES,SESAMO','','','','',601,0,30,0,0,0,0,'',0.33,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'38701.png','','',0,0.000,0),(12136,38710,'QUESO D.O.P SAN SIMON DA','QUESO D.O.P SAN SIMON DA','','','','','','',302,0,36,1,0,0,0,'',10.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'38710.png','','',0,0.000,0),(12137,38744,'PICADA VACUNO BURGER MEAT','PICADA VACUNO BURGER MEAT','','','','','','',201,0,18,1,0,0,0,'',5.10,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'38744.png','','',0,0.000,0),(12138,38767,'PAN BARRA 5 SEMILLAS MASA','PAN BARRA 5 SEMILLAS MASA','','Ingredientes:HARINA DE TRIGO,agua,MASA MADRE NATURAL (TRIGO);Semillas:SESAMO,lino dorado,lino marron,pipas de girasol,trigo sarraceno,HARINA DE CENTENO,sal,levadura;Emulgentes:lecitina de girasol;Puede contener trazas de:SOJA','','','','',601,0,31,0,0,0,0,'',0.79,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'38767.png','','',0,0.000,0),(12139,38768,'PAN BARRA RUSTICA MASA MA','PAN BARRA RUSTICA MASA MA','','Ingredientes:HARINA DE TRIGO,agua,MASA MADRE NATURAL (TRIGO),sal,levadura,AVENA,Malta DE TRIGO,emulgentes,lecitina de girasol;Puede contener trazas de:SOJA,SESAMO','','','','',601,0,32,0,0,0,0,'',0.79,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'38768.png','','',0,0.000,0),(12140,38769,'PAN 6 CEREALES','PAN 6 CEREALES','','Ingredientes:Cereales:HARINA DE TRIGO,mijo,HARINA DE CENTENO,SEMOLA DE TRIGO,harina de arroz,TRIGO,TRIGO SARRACENO,MASA MADRE DE CENTENO,HARINA DE Malta DE TRIGO,agua;Semillas:lino,pipas de girasol,SESAMO,levadura,sal;Puede contener trazas de:SOJA Y DERIVADOS,SESAMO','','','','',601,0,33,0,0,0,0,'',0.69,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'38769.png','','',0,0.000,0),(12141,38770,'PAN CENTENO','PAN CENTENO','','Ingredientes:HARINA DE CENTENO,CENTENO,HARINA DE TRIGO,agua,sal,levadura,GLUTEN DE TRIGO,HARINA DE Malta DE TRIGO;Puede contener trazas de:SOJA,SESAMO','','','','',601,0,34,0,0,0,0,'',0.55,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'38770.png','','',0,0.000,0),(12142,38776,'PAN MINI RUSTICO','PAN MINI RUSTICO','','Ingredientes:HARINA DE TRIGO,agua,levadura,sal;Puede contener trazas de:SESAMO,SOJA','','','','',601,0,35,0,0,0,0,'',0.20,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'38776.png','','',0,0.000,0),(12143,38785,'PAN DE LA ABUELA MASA MAD','PAN DE LA ABUELA MASA MAD','','Ingredientes:HARINA DE TRIGO,agua,MASA MADRE DE TRIGO,sal,levadura;Puede contener trazas de:SOJA,SESAMO,FRUTOS DE CÁSCARA','','','','',601,0,36,0,0,0,0,'',1.10,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'38785.png','','',0,0.000,0),(12144,38841,'PICADA VAC/PORC BURGER ME','PICADA VAC/PORC BURGER ME','','','','','','',208,0,16,1,0,0,0,'',4.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'38841.png','','',0,0.000,0),(12145,38971,'BOMBON ALMENDRADO','BOMBON ALMENDRADO','','','','','','',603,0,32,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38971.png','','',0,0.000,0),(12146,38972,'TRONQUITO DE ALMENDRA','TRONQUITO DE ALMENDRA','','','','','','',603,0,33,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'38972.png','','',0,0.000,0),(12147,39061,'QUESO DECAVA BARUS','QUESO DECAVA BARUS','','','','','','',302,0,37,1,0,0,0,'',14.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39061.png','','',0,0.000,0),(12148,39062,'MANZANA GOLDEN BOLSA 1,5K','MANZANA GOLDEN BOLSA 1,5K','','Variedad:golden;Origen:España;Categoria:primera;Calibre:60mm/+','','','','',501,0,46,0,0,0,0,'',2.00,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39062.png','','',0,0.000,0),(12149,39071,'HAMBURGUESA VACUNO CON QU','HAMBURGUESA VACUNO CON QU','','','','','','',201,0,19,1,0,0,0,'',2.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'39071.png','','',0,0.000,0),(12150,39072,'HAMBURGUESA MINI POLLO BJ','HAMBURGUESA MINI POLLO BJ','','','','','','',205,0,24,1,0,0,0,'',1.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'39072.png','','',0,0.000,0),(12151,39102,'ROLLOS DE ANIS','ROLLOS DE ANIS','','','','','','',602,0,32,0,0,0,0,'',1.49,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'39102.png','','',0,0.000,0),(12152,39121,'CALABAZA PARTIDA','CALABAZA PARTIDA','','Origen:España;Categoria:i;Variedad:cacahuet/tasky delice','','','','',502,0,86,0,0,0,0,'',1.15,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39121.png','','',0,0.000,0),(12153,39123,'NARANJA GRANEL ZUMO','NARANJA GRANEL ZUMO','','Variedad:navelina;Origen:España;Calibre:6/8;Categoria:ii','','','','',501,0,47,0,0,0,0,'',1.30,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39123.png','','',0,0.000,0),(12154,39136,'ESCALIVADA','ESCALIVADA','','','','','','',502,0,87,0,0,0,0,'',1.95,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39136.png','','',0,0.000,0),(12155,39137,'ALCACHOFA MARINADA','ALCACHOFA MARINADA','','','','','','',502,0,88,0,0,0,0,'',1.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39137.png','','',0,0.000,0),(12156,39138,'PIMIENTO ASADO TIRAS','PIMIENTO ASADO TIRAS','','','','','','',502,0,89,0,0,0,0,'',1.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39138.png','','',0,0.000,0),(12157,39199,'PATATA PARA FREIR 3KG','PATATA PARA FREIR 3KG','','','','','','',502,0,90,0,0,0,0,'',3.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39199.png','','',0,0.000,0),(12158,39217,'PAVO NACIONAL +9 KG','PAVO NACIONAL +9 KG','','','','','','',205,0,25,1,0,0,0,'',4.40,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'39217.png','','',0,0.000,0),(12159,39218,'RELLENO MILHOJAS CERDO','RELLENO MILHOJAS CERDO','','','','','','',208,0,17,1,0,0,0,'',9.20,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'39218.png','','',0,0.000,0),(12160,39220,'FILETE MERLUZA GRANEL','FILETE MERLUZA GRANEL','','','','','','',701,0,11,1,0,0,0,'',15.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'39220.png','','',0,0.000,0),(12161,39408,'ATUN AHUMADO','ATUN AHUMADO','','','','','','',301,0,41,1,0,0,0,'',27.50,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39408.png','','',0,0.000,0),(12162,39410,'MOJAMA EXTRA','MOJAMA EXTRA','','','','','','',301,0,42,1,0,0,0,'',41.90,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39410.png','','',0,0.000,0),(12163,39421,'JAMON SERRANO DESH. RESER','JAMON SERRANO DESH. RESER','','Ingredientes:jamon de cerdo,sal,azucar;Conservadores:E-250,E-252;Antioxidante:E-301;Condiciones de conservacion:en sitio fresco y seco','','','','',301,0,43,1,0,0,0,'',15.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39421.png','','',0,0.000,0),(12164,39447,'VERDURA MICROONDAS','VERDURA MICROONDAS','','','','','','',502,0,91,0,0,0,0,'',1.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39447.png','','',0,0.000,0),(12165,39450,'JAMON CURADO DUROC 100%','JAMON CURADO DUROC 100%','','Ingredientes:jamon de cerdo,sal,azucar;Reguladores:E-331(iii);Conservadores:E-250,E-252;Antioxidante:E-301;Condiciones de conservacion:en sitio fresco y seco','','','','',301,0,44,1,0,0,0,'',39.00,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39450.png','','',0,0.000,0),(12166,39451,'JAMON SERRANO RVA. PIEZA','JAMON SERRANO RVA. PIEZA','','','','','','',301,0,45,1,0,0,0,'',48.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39451.png','','',0,0.000,0),(12167,39456,'PANQUEMADO','PANQUEMADO','','','','','','',602,0,33,0,0,0,0,'',2.25,0,0.00,0,0,6,0,'',4,0,0.000,'','-',0,'','-',0,'39456.png','','',0,0.000,0),(12168,39483,'BUTIFARRA NEGRA','BUTIFARRA NEGRA','','','','','','',301,0,46,1,0,0,0,'',7.99,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'39483.png','','',0,0.000,0),(12169,39484,'CABEZA LOMO A LA PIMIENTA','CABEZA LOMO A LA PIMIENTA','','Ingredientes:cabeza de cerdo,sal,especias naturales;Antioxidante:E-301;Conservadores:E-250,E-252;Puede contener trazas de:SOJA Y DERIVADOS;Condiciones de conservacion:sitio fresco y seco','','','','',301,0,47,1,0,0,0,'',15.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39484.png','','',0,0.000,0),(12170,39505,'CALÇOTS 25U.','CALÇOTS 25U.','','','','','','',502,0,92,0,0,0,0,'',4.25,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'39505.png','','',0,0.000,0),(12171,39568,'PAN TIGRE MEDIO CORTADO','PAN TIGRE MEDIO CORTADO','','','','','','',601,0,37,0,0,0,0,'',1.69,0,0.00,0,0,6,0,'',5,0,0.000,'','-',0,'','-',0,'39568.png','','',0,0.000,0),(12172,39569,'PAN INTEGRAL MEDIO CORTAD','PAN INTEGRAL MEDIO CORTAD','','','','','','',601,0,38,0,0,0,0,'',1.45,0,0.00,0,0,6,0,'',5,0,0.000,'','-',0,'','-',0,'39569.png','','',0,0.000,0),(12173,39570,'PAN 6 CEREALES MEDI CORTA','PAN 6 CEREALES MEDI CORTA','','','','','','',601,0,39,0,0,0,0,'',1.45,0,0.00,0,0,6,0,'',5,0,0.000,'','-',0,'','-',0,'39570.png','','',0,0.000,0),(12174,39571,'PAN MULTICEREA MEDIO NEGR','PAN MULTICEREA MEDIO NEGR','','','','','','',601,0,40,0,0,0,0,'',1.55,0,0.00,0,0,6,0,'',5,0,0.000,'','-',0,'','-',0,'39571.png','','',0,0.000,0),(12175,39572,'PAN DE PASAS','PAN DE PASAS','','','','','','',601,0,41,0,0,0,0,'',2.40,0,0.00,0,0,6,0,'',5,0,0.000,'','-',0,'','-',0,'39572.png','','',0,0.000,0),(12176,39638,'SALCHICHON DE PAVO','SALCHICHON DE PAVO','','','','','','',301,0,48,1,0,0,0,'',10.75,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39638.png','','',0,0.000,0),(12177,39639,'CHORIZO DE PAVO','CHORIZO DE PAVO','','','','','','',301,0,49,1,0,0,0,'',10.75,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39639.png','','',0,0.000,0),(12178,39640,'PAVO TRUFADO','PAVO TRUFADO','','','','','','',301,0,50,1,0,0,0,'',7.85,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'39640.png','','',0,0.000,0),(12179,39681,'SALCHICHON DE LEON','SALCHICHON DE LEON','','Ingredientes:magro de cerdo,panceta de cerdo,dextrina,LECHE EN POLVO,dextrosa,especias;Conservadores:E-250,E-252;Antioxidante:E-301;Colorantes:E-120;Condiciones de conservacion:sitio fresco y seco','','','','',301,0,51,1,0,0,0,'',12.55,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39681.png','','',0,0.000,0),(12180,39682,'CHORIZO DE LEON','CHORIZO DE LEON','','Ingredientes:magro de cerdo,panceta de cerdo,sal,dextrina,PROTEINAS LACTEAS,ajo,dextrosa,azucar;Antioxidante:E-331(iii),E-301;Conservadores:E-252;Condiciones de conservacion:sitio fresco y seco','','','','',301,0,52,1,0,0,0,'',12.55,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39682.png','','',0,0.000,0),(12181,39688,'QUESO MEZCLA IBERICO GRAN','QUESO MEZCLA IBERICO GRAN','','Ingredientes:LECHE CRUDA DE OVEJA,LECHE DE VACA,LECHE DE CABRA,sal,fermentos lacticos,cuajo,cloruro calcico,LISOZIMA (HUEVO);Condiciones de conservacion:sitio fresco y seco','','','','',302,0,38,1,0,0,0,'',12.90,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39688.png','','',0,0.000,0),(12182,39689,'QUESO DE OVEJA AÑEJO','QUESO DE OVEJA AÑEJO','','','','','','',302,0,39,1,0,0,0,'',13.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39689.png','','',0,0.000,0),(12183,39690,'QUESO DE CABRA SEMICURADO','QUESO DE CABRA SEMICURADO','','','','','','',302,0,40,1,0,0,0,'',13.90,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39690.png','','',0,0.000,0),(12184,39738,'CHULETON TERNERA ROSADA','CHULETON TERNERA ROSADA','','','','','','',201,0,20,1,0,0,0,'',12.90,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'39738.png','','',0,0.000,0),(12185,39771,'QUESO BARRA TIERNO LIGHT','QUESO BARRA TIERNO LIGHT','','','','','','',302,0,41,1,0,0,0,'',9.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'39771.png','','',0,0.000,0),(12186,39849,'HAMBURGUESAS DE CERDO IBÉ','HAMBURGUESAS DE CERDO IBÉ','','','','','','',202,0,26,1,0,0,0,'',2.69,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'39849.png','','',0,0.000,0),(12187,39854,'NAPOLITANA JAMON Y QUESO','NAPOLITANA JAMON Y QUESO','','Ingredientes:Masa:HARINA DE TRIGO;Margarina:grasa de palma,agua,aceite de soja,aceite de girasol;Emulgentes:E-471,sal;Acidulantes:E-330,aromas;Antioxidante:E-306,E-304;Colorantes:E-160(a),azucar,levadura,GLUTEN DE TRIGO;Emulgentes:E-472e;Relleno:QUESO,SUERO LACTEO;Almidón modificado:E-1414,QUESO AZUL;Estabilizantes:E-401,E-415,E-412,PROTEINAS LACTEAS,almidon de maiz,jamon cocido,magro de cerdo,fecula,dextrosa,PROTEINA DE SOJA;Estabilizadores:E-451,E-407,E-410;Antioxidante:E-301,aroma carne,especias;Conservadores:E-250,colorante,E-120;Decoracion:HUEVO LíQUIDO;Puede contener trazas de:PESCADO,FRUTOS DE CÁSCARA,SESAMO','','','','',602,0,34,0,0,0,0,'',0.50,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'39854.png','','',0,0.000,0),(12188,39855,'EMPANADA DE ATUN MINI','EMPANADA DE ATUN MINI','','Ingredientes:Masa:HARINA DE TRIGO,agua,aceite de girasol,vino blanco;Manteca de cerdo:grasa de cerdo;Estabilizadores:E-471,sal,LECHE EN POLVO,LECITINA DE SOJA;Relleno:Pisto:cebolla,pimiento rojo,pimiento verde,aceite de oliva,tomate,azucar;Almidón modificado:E-1420,sal;Estabilizadores:E-412;Acidificantes:E-330,ATUN,aceite de girasol,sal,PAN RALLADO,caldo de carne,pimenton;Cobertura:HUEVO LíQUIDO;Puede contener trazas de:FRUTOS DE CÁSCARA,APIO Y DERIVADOS,SESAMO','','','','',602,0,35,0,0,0,0,'',0.95,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'39855.png','','',0,0.000,0),(12189,39856,'SUSO DE CREMA','SUSO DE CREMA','','Ingredientes:Masa:HARINA DE TRIGO,agua,grasa de palma,aceite de girasol,aceite de coco,dextrosa,levadura,HARINA DE SOJA,SUERO DE LECHE EN POLVO;Gasificantes:E-450,E-500,sal;Emulgentes:E-471,E-481,LECHE DES. EN POLVO,HUEVO ENTERO EN POLVO;Colorantes:E-160(a),aroma;Relleno:agua,jarabe de glucosa-fructosa,azucar;Almidón modificado:E-1442,grasa de palma,ALMIDON DE TRIGO,sal;Acidulantes:E-330,aroma;Emulgentes:E-435;Colorante natural:E-171,E-160(a);Puede contener trazas de:FRUTOS DE CÁSCARA,producto descongelado','','','','',602,0,36,0,0,0,0,'',0.75,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'39856.png','','',0,0.000,0),(12190,39976,'TARTA QUESO OREO','TARTA QUESO OREO','','','','','','',603,0,34,0,0,0,0,'',5.50,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'39976.png','','',0,0.000,0),(12191,40023,'SALCHICHON ANCHETE PIMIEN','SALCHICHON ANCHETE PIMIEN','','','','','','',301,0,53,1,0,0,0,'',10.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'40023.png','','',0,0.000,0),(12192,40074,'ENSALADA NEW YORK','ENSALADA NEW YORK','','','','','','',502,0,93,0,0,0,0,'',2.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40074.png','','',0,0.000,0),(12193,40150,'MAGDALENA CASERA CHOCO. 6','MAGDALENA CASERA CHOCO. 6','','','','','','',602,0,37,0,0,0,0,'',1.69,0,0.00,0,0,6,0,'',21,0,0.000,'','-',0,'','-',0,'40150.png','','',0,0.000,0),(12194,40151,'MAGDALENA CASERA 6U','MAGDALENA CASERA 6U','','','','','','',602,0,38,0,0,0,0,'',1.69,0,0.00,0,0,6,0,'',21,0,0.000,'','-',0,'','-',0,'40151.png','','',0,0.000,0),(12195,40152,'MAGDALENA INTEGRAL S/AZ 6','MAGDALENA INTEGRAL S/AZ 6','','','','','','',602,0,39,0,0,0,0,'',2.00,0,0.00,0,0,6,0,'',15,0,0.000,'','-',0,'','-',0,'40152.png','','',0,0.000,0),(12196,40153,'MAPLE PECAN','MAPLE PECAN','','Ingredientes:Masa:HARINA DE TRIGO,margarina,grasa de palma,ACEITE DE SOJA,aceite de girasol,jarabe de glucosa-fructosa,agua;Emulgentes:E-322,E-471;Acidulantes:E-330,aromas;Colorantes:E-160a,agua,azucar,HUEVO LIQUIDO PASTEURIZADO,levadura,LECHE DES. EN POLVO,sal,GLUTEN;Relleno:azucar,agua;Almidón modificado:E-1442,LECHE CONDENSADA,aroma;Gelificantes:E-460,E-466,jarabe de azucar invertido;Acidulantes:E-270;Espesantes:E-415,azucar;Cobertura:NUEZ PECAN,HUEVO LIQUIDO PASTEURIZADO,estabilizador,E-440,regulador de acidez,E-450,conservantes,E-202;Puede contener trazas de:SOJA Y DERIVADOS,SESAMO','','','','',602,0,40,0,0,0,0,'',0.70,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'40153.png','','',0,0.000,0),(12197,40268,'ESPARRAGO VERDE GRUESO MA','ESPARRAGO VERDE GRUESO MA','','','','','','',502,0,94,0,0,0,0,'',2.75,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40268.png','','',0,0.000,0),(12198,40320,'CUARTOS TRASEROS POLLO RU','CUARTOS TRASEROS POLLO RU','','','','','','',205,0,26,1,0,0,0,'',3.49,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'40320.png','','',0,0.000,0),(12199,40322,'CUARTOS TRASE. POLLO RURA','CUARTOS TRASE. POLLO RURA','','','','','','',205,0,27,1,0,0,0,'',3.45,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'40322.png','','',0,0.000,0),(12200,40350,'ENSALADA CAPRESE','ENSALADA CAPRESE','','','','','','',502,0,95,0,0,0,0,'',2.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40350.png','','',0,0.000,0),(12201,40385,'CODILLO CERDO ASADO','CODILLO CERDO ASADO','','','','','','',208,0,18,1,0,0,0,'',4.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'40385.png','','',0,0.000,0),(12202,40447,'HOGAZA INTEGRAL MULTIGRAI','HOGAZA INTEGRAL MULTIGRAI','','Ingredientes:HARINA DE TRIGO,agua,MASA MADRE DE TRIGO,TRIGO,MASA MADRE DE CENTENO,mijo,CEBADA,CENTENO,AVENA,TRIGO SARRACENO,sal,teff,harina de quinoa,sorgo,amaranto,agua,SEMOLA DE TRIGO,sal,levadura;Puede contener trazas de:SOJA,FRUTOS DE CÁSCARA,SESAMO','','','','',601,0,42,0,0,0,0,'',1.60,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'40447.png','','',0,0.000,0),(12203,40453,'HOGAZA PASAS Y NUECES','HOGAZA PASAS Y NUECES','','','','','','',602,0,41,0,0,0,0,'',2.69,0,0.00,0,0,6,0,'',14,0,0.000,'','-',0,'','-',0,'40453.png','','',0,0.000,0),(12204,40459,'MANZANA ROYAL GALA SELECC','MANZANA ROYAL GALA SELECC','','Variedad:royal gala;Origen:Italia;Categoria:i;Calibre:80/85','','','','',501,0,12,1,0,0,0,'',2.25,0,0.00,1,45,4,0,'',0,0,0.000,'','-',0,'','-',0,'40459.png','','',0,0.000,0),(12205,40495,'(GRANEL) BRACITO DE FRESA','(GRANEL) BRACITO DE FRESA','','','','','','',603,0,35,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'40495.png','','',0,0.000,0),(12206,40496,'(GRANEL) BORRACHO AL CACA','(GRANEL) BORRACHO AL CACA','','','','','','',603,0,36,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'40496.png','','',0,0.000,0),(12207,40498,'(GRANEL) CAÑA CACAO','(GRANEL) CAÑA CACAO','','','','','','',603,0,37,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'40498.png','','',0,0.000,0),(12208,40499,'(GRANEL) BIZCOCHO BAÑADO','(GRANEL) BIZCOCHO BAÑADO','','','','','','',603,0,38,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'40499.png','','',0,0.000,0),(12209,40500,'(GRANEL) MIGUELITO CACAO','(GRANEL) MIGUELITO CACAO','','','','','','',603,0,39,0,0,0,0,'',3.95,0,0.00,1,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'40500.png','','',0,0.000,0),(12210,40558,'ROSEGON','ROSEGON','','Ingredientes:HARINA DE TRIGO,azucar,ALMENDRAS,(12%),HUEVO,aceite de girasol,EXTRACTO DE Malta DE CEBADA,gasificantes:,E-503,E-500,aromas,limon,almendras,vainilla;Puede contener trazas de:LECHE,SOJA,FRUTOS DE CÁSCARA,CACAHUETES','','','','',602,0,42,0,0,0,0,'',8.25,0,0.00,1,0,6,0,'',30,0,0.000,'','-',0,'','-',0,'40558.png','','',0,0.000,0),(12211,40580,'COCO NATURAL','COCO NATURAL','','','','','','',501,0,49,0,0,0,0,'',1.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40580.png','','',0,0.000,0),(12212,40581,'PANECILLO HOLANDES BLANCO','PANECILLO HOLANDES BLANCO','','','','','','',601,0,43,0,0,0,0,'',1.59,0,0.00,0,0,6,0,'',5,0,0.000,'','-',0,'','-',0,'40581.png','','',0,0.000,0),(12213,40582,'PANECILLO HOLANDES INTEGR','PANECILLO HOLANDES INTEGR','','','','','','',601,0,44,0,0,0,0,'',1.59,0,0.00,0,0,6,0,'',5,0,0.000,'','-',0,'','-',0,'40582.png','','',0,0.000,0),(12214,40651,'AJOS NEGROS','AJOS NEGROS','','','','','','',502,0,96,0,0,0,0,'',5.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40651.png','','',0,0.000,0),(12215,40653,'QUESO MEZCLA SEMI MEDIA P','QUESO MEZCLA SEMI MEDIA P','','','','','','',302,0,42,1,0,0,0,'',9.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'40653.png','','',0,0.000,0),(12216,40664,'BROCHETAS PRIMAVERA MINI','BROCHETAS PRIMAVERA MINI','','','','','','',208,0,19,1,0,0,0,'',2.65,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'40664.png','','',0,0.000,0),(12217,40665,'BROCHETA GRILL MINI BJA.','BROCHETA GRILL MINI BJA.','','','','','','',208,0,20,1,0,0,0,'',2.65,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'40665.png','','',0,0.000,0),(12218,40675,'FLORETES DE COLIFLOR','FLORETES DE COLIFLOR','','','','','','',502,0,97,0,0,0,0,'',1.45,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40675.png','','',0,0.000,0),(12219,40676,'FLORETES DE BROCOLI','FLORETES DE BROCOLI','','','','','','',502,0,98,0,0,0,0,'',1.45,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40676.png','','',0,0.000,0),(12220,40678,'MAIZ EN GRANO','MAIZ EN GRANO','','','','','','',502,0,99,0,0,0,0,'',1.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40678.png','','',0,0.000,0),(12221,40707,'PAN BOCADILLO 5U','PAN BOCADILLO 5U','','Ingredientes:HARINA DE TRIGO,agua,levadura,sal,emulgentes,E-472e','','','','',601,0,45,0,0,0,0,'',1.10,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'40707.png','','',0,0.000,0),(12222,40709,'SOBRASADA CULAR','SOBRASADA CULAR','','','','','','',301,0,54,1,0,0,0,'',8.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'40709.png','','',0,0.000,0),(12223,40733,'LONGANIZA DE PASCUA','LONGANIZA DE PASCUA','','Ingredientes:carne de cerdo,sal,almidon de maiz,PROTEINA DE SOJA,especias,dextrosa,azucar,LACTOSA;Estabilizadores:E-451,E-452;Antioxidante:E-331,E-301,fermentos lacticos,aroma;Conservadores:E-250,E-252;Colorante natural:E-120;Condiciones de conservacion:sitio fresco y seco','','','','',301,0,55,1,0,0,0,'',10.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'40733.png','','',0,0.000,0),(12224,40736,'ZANAHORIA BABY','ZANAHORIA BABY','','','','','','',502,0,100,0,0,0,0,'',1.05,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40736.png','','',0,0.000,0),(12225,40738,'CHAMPIÑON PORTOBELLO','CHAMPIÑON PORTOBELLO','','','','','','',502,0,101,0,0,0,0,'',1.25,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40738.png','','',0,0.000,0),(12226,40741,'QUESO DE CABRA DELICIAS','QUESO DE CABRA DELICIAS','','','','','','',302,0,43,1,0,0,0,'',14.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'40741.png','','',0,0.000,0),(12227,40802,'SALAMI EXTRA','SALAMI EXTRA','','','','','','',301,0,56,1,0,0,0,'',6.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'40802.png','','',0,0.000,0),(12228,40817,'MANZANA FUJI PANDA','MANZANA FUJI PANDA','','Variedad:fuji;Origen:China;Calibre:78/80;Categoria:i','','','','',501,0,28,1,0,0,0,'',2.99,0,0.00,1,93,4,0,'',0,0,0.000,'','-',0,'','-',0,'40817.png','','',0,0.000,0),(12229,40877,'QUESO RULO CABRA NATURAL','QUESO RULO CABRA NATURAL','','','','','','',302,0,44,1,0,0,0,'',12.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'40877.png','','',0,0.000,0),(12230,40886,'COLIFLOR UNIDAD','COLIFLOR UNIDAD','','','','','','',502,0,102,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'40886.png','','',0,0.000,0),(12231,41039,'ALBAHACA FRESCA','ALBAHACA FRESCA','','','','','','',502,0,103,0,0,0,0,'',1.15,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41039.png','','',0,0.000,0),(12232,41062,'ENTRECOT VACUNO MADURADO','ENTRECOT VACUNO MADURADO','','','','','','',201,0,21,1,0,0,0,'',15.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41062.png','','',0,0.000,0),(12233,41122,'PALETA CURADA DUROC PIEZA','PALETA CURADA DUROC PIEZA','','','','','','',301,0,57,1,0,0,0,'',23.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41122.png','','',0,0.000,0),(12234,41156,'FILETE DE POTRO BANDEJA','FILETE DE POTRO BANDEJA','','','','','','',204,0,1,1,0,0,0,'',17.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41156.png','','',0,0.000,0),(12235,41168,'CEBOLLINO FRESCO','CEBOLLINO FRESCO','','','','','','',502,0,104,0,0,0,0,'',1.15,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41168.png','','',0,0.000,0),(12236,41169,'CILANTRO FRESCO','CILANTRO FRESCO','','','','','','',502,0,105,0,0,0,0,'',1.15,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41169.png','','',0,0.000,0),(12237,41170,'HIERBABUENA FRESCA','HIERBABUENA FRESCA','','','','','','',502,0,106,0,0,0,0,'',1.15,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41170.png','','',0,0.000,0),(12238,41196,'PAN BARRA GALLEGA 2 UNIDA','PAN BARRA GALLEGA 2 UNIDA','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura;Emulgentes:E-472e','','','','',601,0,46,0,0,0,0,'',1.26,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'41196.png','','',0,0.000,0),(12239,41198,'PECHUGA PAVO ARTESANA','PECHUGA PAVO ARTESANA','','','','','','',301,0,58,1,0,0,0,'',12.95,0,0.00,1,0,3,0,'',15,0,0.000,'','-',0,'','-',0,'41198.png','','',0,0.000,0),(12240,41206,'ENSALADA COMPLETA OPTIMA','ENSALADA COMPLETA OPTIMA','','','','','','',502,0,107,0,0,0,0,'',2.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41206.png','','',0,0.000,0),(12241,41224,'QUESO CURADO HOLANDES P.','QUESO CURADO HOLANDES P.','','','','','','',302,0,45,1,0,0,0,'',11.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41224.png','','',0,0.000,0),(12242,41225,'QUESO BOLA TIERNO HOLANDE','QUESO BOLA TIERNO HOLANDE','','','','','','',302,0,46,1,0,0,0,'',7.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41225.png','','',0,0.000,0),(12243,41226,'JAMON RESERVA SELECCION P','JAMON RESERVA SELECCION P','','','','','','',301,0,59,1,0,0,0,'',59.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41226.png','','',0,0.000,0),(12244,41235,'PECHUGA FILETES POLLO RUR','PECHUGA FILETES POLLO RUR','','','','','','',205,0,28,1,0,0,0,'',7.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41235.png','','',0,0.000,0),(12245,41291,'ENSAIMADA CABELLO ANGEL','ENSAIMADA CABELLO ANGEL','','','','','','',602,0,43,0,0,0,0,'',2.75,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'41291.png','','',0,0.000,0),(12246,41297,'ZANAHORIA ECOLOGICA','ZANAHORIA ECOLOGICA','','','','','','',502,0,108,0,0,0,0,'',1.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41297.png','','',0,0.000,0),(12247,41298,'PUERROS ECOLOGICOS','PUERROS ECOLOGICOS','','','','','','',502,0,109,0,0,0,0,'',2.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41298.png','','',0,0.000,0),(12248,41299,'PATATAS ECOLOGICAS 2KG','PATATAS ECOLOGICAS 2KG','','','','','','',502,0,110,0,0,0,0,'',3.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41299.png','','',0,0.000,0),(12249,41300,'CEBOLLA ECOLOGICA','CEBOLLA ECOLOGICA','','','','','','',502,0,111,0,0,0,0,'',1.69,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41300.png','','',0,0.000,0),(12250,41302,'TOMATE CHERRY ECOLOGICO','TOMATE CHERRY ECOLOGICO','','','','','','',502,0,112,0,0,0,0,'',1.50,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41302.png','','',0,0.000,0),(12251,41304,'TOMATE RAMA ECOLOGICO','TOMATE RAMA ECOLOGICO','','','','','','',502,0,113,0,0,0,0,'',1.70,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41304.png','','',0,0.000,0),(12252,41305,'AJOS SECOS ECOLOGICOS','AJOS SECOS ECOLOGICOS','','','','','','',502,0,114,0,0,0,0,'',1.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41305.png','','',0,0.000,0),(12253,41308,'PAN INTEGRAL BAGUETTINA 4','PAN INTEGRAL BAGUETTINA 4','','Ingredientes:HARINA DE TRIGO,agua,SALVADO DE TRIGO,levadura,sal;Emulgentes:E-472;Antioxidante:E-300,E-322 (SOJA);Puede contener trazas de:LECHE Y DERIVADOS,SOJA Y DERIVADOS,SESAMO','','','','',601,0,47,0,0,0,0,'',1.00,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'41308.png','','',0,0.000,0),(12254,41318,'PAN BARRA INTEGRAL 2U','PAN BARRA INTEGRAL 2U','','Ingredientes:HARINA DE TRIGO,agua,SALVADO DE TRIGO,sal,levadura;Emulgentes:E-472e;Puede contener trazas de:LECHE,SOJA,SESAMO,FRUTOS DE CÁSCARA','','','','',601,0,48,0,0,0,0,'',1.10,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'41318.png','','',0,0.000,0),(12255,41326,'QUESO GOUDA TIERNO','QUESO GOUDA TIERNO','','','','','','',302,0,47,1,0,0,0,'',8.50,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41326.png','','',0,0.000,0),(12256,41327,'QUESO GOUDA DE CABRA','QUESO GOUDA DE CABRA','','','','','','',302,0,48,1,0,0,0,'',12.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41327.png','','',0,0.000,0),(12257,41328,'QUESO GOUDA COMINO','QUESO GOUDA COMINO','','','','','','',302,0,49,1,0,0,0,'',8.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41328.png','','',0,0.000,0),(12258,41329,'QUESO BOLA SEMI HOLANDES','QUESO BOLA SEMI HOLANDES','','','','','','',302,0,50,1,0,0,0,'',8.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41329.png','','',0,0.000,0),(12259,41331,'BUTIFARRA BLANCA','BUTIFARRA BLANCA','','','','','','',301,0,60,1,0,0,0,'',7.95,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'41331.png','','',0,0.000,0),(12260,41338,'CATALANA COCIDA','CATALANA COCIDA','','','','','','',301,0,61,1,0,0,0,'',7.95,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'41338.png','','',0,0.000,0),(12261,41345,'HAMBURGUESA DE NOVILLA 2U','HAMBURGUESA DE NOVILLA 2U','','','','','','',201,0,22,1,0,0,0,'',3.49,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41345.png','','',0,0.000,0),(12262,41368,'QUESO STILTON DOP','QUESO STILTON DOP','','','','','','',302,0,51,1,0,0,0,'',17.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41368.png','','',0,0.000,0),(12263,41369,'LLONGANISSA AL PIMENTÓ EX','LLONGANISSA AL PIMENTÓ EX','','Ingredientes:carne de cerdo,dextrina,sal,pimenton,azucar,especias,ajo;Conservadores:E-252,dextrosa;Antioxidantes:E-301,E-331,aroma','','','','',301,0,62,1,0,0,0,'',13.99,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'41369.png','','',0,0.000,0),(12264,41389,'KIWI GREEN ECOLOGICO','KIWI GREEN ECOLOGICO','','','','','','',501,0,51,0,0,0,0,'',3.80,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41389.png','','',0,0.000,0),(12265,41390,'PEPINO HOLANDES ECOLOGICO','PEPINO HOLANDES ECOLOGICO','','','','','','',502,0,115,0,0,0,0,'',1.25,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41390.png','','',0,0.000,0),(12266,41436,'ENSALADA CESAR','ENSALADA CESAR','','','','','','',502,0,116,0,0,0,0,'',2.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41436.png','','',0,0.000,0),(12267,41497,'JAMON BELLOTA IBERICO PIE','JAMON BELLOTA IBERICO PIE','','','','','','',301,0,63,1,0,0,0,'',295.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41497.png','','',0,0.000,0),(12268,41500,'LOMO FRESCO ADOBADO SKIN','LOMO FRESCO ADOBADO SKIN','','','','','','',208,0,21,1,0,0,0,'',3.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41500.png','','',0,0.000,0),(12269,41564,'PAIN CHOC','PAIN CHOC','','Ingredientes:Masa:HARINA DE TRIGO,MANTEQUILLA (LECHE),agua,azucar,levadura,GLUTEN,sal,fibra vegetal,grasa de palma;Emulgentes:E-471;Relleno:azucar,pasta de cacao,manteca de cacao;Emulgentes:E-322 (SOJA);Cobertura:HUEVO LIQUIDO PASTEURIZADO,dextrosa;Puede contener trazas de:FRUTOS DE CÁSCARA,MOSTAZA','','','','',602,0,44,0,0,0,0,'',0.60,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'41564.png','','',0,0.000,0),(12270,41566,'PAÑUELO CARBONARA','PAÑUELO CARBONARA','','Ingredientes:Masa:HARINA DE TRIGO,margarina,agua,azucar;Emulgentes:E-471,E-472e;Espesantes:E-415,E-466,E-412,dextrosa,enzimas;Agente de tratamiento de la harina:E-920,levadura,sal;Relleno:almidon de patata,SUERO LACTEO,PROTEINAS LACTEAS,aceite de palma,HUEVO EN POLVO,ALMIDON DE TRIGO,sal,QUESO,jarabe de glucosa,cebolla,PROTEINA DE SOJA,aromas naturales,especias;Estabilizantes:E-412,aroma humo,bacon;Puede contener trazas de:PESCADO Y DERIVADOS,FRUTOS DE CÁSCARA,APIO Y DERIVADOS,SESAMO,SULFITOS','','','','',602,0,45,0,0,0,0,'',1.10,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'41566.png','','',0,0.000,0),(12271,41568,'BAGUETTE ANTIGA','BAGUETTE ANTIGA','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura;Emulgentes:E-472e','','','','',601,0,49,0,0,0,0,'',0.70,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'41568.png','','',0,0.000,0),(12272,41580,'FILETE TIERNO BLACK ANGUS','FILETE TIERNO BLACK ANGUS','','','','','','',201,0,23,1,0,0,0,'',14.95,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41580.png','','',0,0.000,0),(12273,41597,'SALCHICHON IBERICO','SALCHICHON IBERICO','','Ingredientes:carne de cerdo ibérico,grasa de cerdo ibérico,LACTOSA,PROTEINA DE SOJA,dextrosa,dextrina;Emulgentes:E-451i;Conservadores:E-250,E-252;Antioxidante:E-316,E-331iiii;Potenciador del sabor:E-621;Colorante natural:E-120;Condiciones de conservacion:conservar en sitio fresco yseco','','','','',301,0,64,1,0,0,0,'',16.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41597.png','','',0,0.000,0),(12274,41620,'ENSAIMADA CHOCOLATE','ENSAIMADA CHOCOLATE','','','','','','',602,0,46,0,0,0,0,'',2.75,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'41620.png','','',0,0.000,0),(12275,41671,'TARTA FRUTOS SECOS','TARTA FRUTOS SECOS','','','','','','',603,0,40,0,0,0,0,'',1.85,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'41671.png','','',0,0.000,0),(12276,41672,'TARTA MANZANA','TARTA MANZANA','','','','','','',603,0,41,0,0,0,0,'',1.85,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'41672.png','','',0,0.000,0),(12277,41684,'COMPAGNO','COMPAGNO','','','','','','',210,0,15,1,0,0,0,'',3.25,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41684.png','','',0,0.000,0),(12278,41692,'JAMON CURADO GRAN RVA. PI','JAMON CURADO GRAN RVA. PI','','','','','','',301,0,65,1,0,0,0,'',69.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41692.png','','',0,0.000,0),(12279,41693,'JAMON DUROC DESH. GRAN RV','JAMON DUROC DESH. GRAN RV','','','','','','',301,0,66,1,0,0,0,'',22.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41693.png','','',0,0.000,0),(12280,41694,'LOMO IBERICO','LOMO IBERICO','','Ingredientes:lomo de cerdo iberico,sal,pimenton,ajo,oregano,LACTOSA,dextrosa,azucar,PROTEINAS LACTEAS;Conservadores:E-250,E-252;Antioxidante:E-301,aceite de oliva;Condiciones de conservacion:conservar en sitio fresco y seco','','','','',301,0,67,1,0,0,0,'',35.00,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41694.png','','',0,0.000,0),(12281,41695,'QUESO RONCAL D.O.','QUESO RONCAL D.O.','','','','','','',302,0,52,1,0,0,0,'',19.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41695.png','','',0,0.000,0),(12282,41696,'QUESO IDIAZABAL AHUMADO D','QUESO IDIAZABAL AHUMADO D','','','','','','',302,0,53,1,0,0,0,'',19.50,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'41696.png','','',0,0.000,0),(12283,41712,'CONEJO TROCEADO BANDEJA','CONEJO TROCEADO BANDEJA','','','','','','',205,0,29,1,0,0,0,'',2.00,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41712.png','','',0,0.000,0),(12284,41755,'TARTA MOUSSE TRES CHOCOLA','TARTA MOUSSE TRES CHOCOLA','','','','','','',603,0,42,0,0,0,0,'',7.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'41755.png','','',0,0.000,0),(12285,41756,'TARTA QUESO Y FRESAS','TARTA QUESO Y FRESAS','','','','','','',603,0,43,0,0,0,0,'',5.49,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'41756.png','','',0,0.000,0),(12286,41762,'TRUCHA 500GR UNIDAD','TRUCHA 500GR UNIDAD','','','','','','',701,0,12,1,0,0,0,'',3.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'41762.png','','',0,0.000,0),(12287,41768,'LOMO CERDO IBERICO','LOMO CERDO IBERICO','','','','','','',202,0,27,1,0,0,0,'',9.95,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41768.png','','',0,0.000,0),(12288,41836,'POLLO RELLENO ARANDANOS Q','POLLO RELLENO ARANDANOS Q','','','','','','',301,0,68,1,0,0,0,'',14.95,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'41836.png','','',0,0.000,0),(12289,41844,'SEPIA SUCIA FRESCA','SEPIA SUCIA FRESCA','','','','','','',702,0,18,1,0,0,0,'',13.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'41844.png','','',0,0.000,0),(12290,41861,'HIGADOS DE CONEJO BANDEJA','HIGADOS DE CONEJO BANDEJA','','','','','','',205,0,30,1,0,0,0,'',7.90,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41861.png','','',0,0.000,0),(12291,41880,'BROCOLI ECOLOGICO','BROCOLI ECOLOGICO','','','','','','',502,0,117,0,0,0,0,'',1.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41880.png','','',0,0.000,0),(12292,41881,'PLATANO ECOLOGICO','PLATANO ECOLOGICO','','','','','','',501,0,52,0,0,0,0,'',2.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41881.png','','',0,0.000,0),(12293,41891,'HAMBURGUESA LABARI BANDEJ','HAMBURGUESA LABARI BANDEJ','','','','','','',201,0,24,1,0,0,0,'',3.10,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'41891.png','','',0,0.000,0),(12294,41893,'ENSALADA LAZOS & RUCULA','ENSALADA LAZOS & RUCULA','','','','','','',502,0,118,0,0,0,0,'',2.95,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'41893.png','','',0,0.000,0),(12295,42023,'PATE OCA TRUFADA PEDRO XI','PATE OCA TRUFADA PEDRO XI','','','','','','',301,0,69,1,0,0,0,'',14.95,0,0.00,1,0,3,0,'',10,0,0.000,'','-',0,'','-',0,'42023.png','','',0,0.000,0),(12296,42029,'PECHUGA PAVO ADOBADA BIEN','PECHUGA PAVO ADOBADA BIEN','','','','','','',208,0,22,1,0,0,0,'',3.10,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'42029.png','','',0,0.000,0),(12297,42030,'ESCALOPIN LOMO EXTRATIERN','ESCALOPIN LOMO EXTRATIERN','','','','','','',202,0,28,1,0,0,0,'',4.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'42030.png','','',0,0.000,0),(12298,42066,'UVA BLANCA SIN SEMILLAS','UVA BLANCA SIN SEMILLAS','','Variedad:early sweet;Origen:sudafrica;Calibre:l;Categoria:i','','','','',501,0,9,1,0,0,0,'',4.79,0,0.00,1,18,4,0,'',0,0,0.000,'','-',0,'','-',0,'42066.png','','',0,0.000,0),(12299,42090,'CROISSANT MANTEQUILLA','CROISSANT MANTEQUILLA','','Ingredientes:Masa:HARINA DE TRIGO,MANTEQUILLA (LECHE),25%,agua,azucar,levadura,GLUTEN DE TRIGO,sal,FIBRA DE TRIGO,grasa de palma;Cobertura:1,HUEVO LIQUIDO PASTEURIZADO,dextrosa;Puede contener trazas de:SOJA,FRUTOS DE CÁSCARA,MOSTAZA','','','','',602,0,47,0,0,0,0,'',0.60,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'42090.png','','',0,0.000,0),(12300,42091,'CROISS MANTEQ CHOCO','CROISS MANTEQ CHOCO','','Ingredientes:Masa:HARINA DE TRIGO,agua,MANTEQUILLA (LECHE),azucar,levadura,emulgentes,E-472e,sal,espesante,E-466,FIBRA DE TRIGO;Relleno:19,azucar,grasa de palma,aceite de girasol,cacao en polvo desgrasado,2,almidon modificado,PASTA DE AVELLANAS,emulgentes,lecitina (girasol),aroma,aroma de vainilla;Cobertura:3,copos de chocolate,2,azucar,pasta de cacao,1,manteca de cacao,0,LECHE ENTERA,grasa anhidrica lactea,aroma de vainilla,HUEVO LIQUIDO PASTEURIZADO,dextrosa;Puede contener trazas de:SOJA,MOSTAZA,SESAMO','','','','',602,0,48,0,0,0,0,'',0.60,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'42091.png','','',0,0.000,0),(12301,42098,'SALMON PORCIONES','SALMON PORCIONES','','','','','','',701,0,13,1,0,0,0,'',4.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'42098.png','','',0,0.000,0),(12302,42099,'LIMA BANDEJA','LIMA BANDEJA','','','','','','',501,0,54,0,0,0,0,'',1.25,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42099.png','','',0,0.000,0),(12303,42108,'CHULETA LOMO CERDO IBERIC','CHULETA LOMO CERDO IBERIC','','','','','','',202,0,29,1,0,0,0,'',7.90,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'42108.png','','',0,0.000,0),(12304,42110,'MORCILLA ASTURIANA','MORCILLA ASTURIANA','','','','','','',210,0,16,1,0,0,0,'',1.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'42110.png','','',0,0.000,0),(12305,42174,'ALMENDRAS CRUDAS PIEL','ALMENDRAS CRUDAS PIEL','','Origen:España;Variedad:largueta;Categoria:i;Calibre:12/14;Puede contener trazas de:CACAHUETES,FRUTOS DE CÁSCARA','','','','',505,0,4,0,0,0,0,'',16.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42174.png','','',0,0.000,0),(12306,42175,'ANACARDOS CRUDOS','ANACARDOS CRUDOS','','Origen:India;Categoria:i;Calibre:w240;Puede contener trazas de:CACAHUETES,FRUTOS DE CÁSCARA','','','','',505,0,5,0,0,0,0,'',16.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42175.png','','',0,0.000,0),(12307,42176,'AVELLANAS CRUDAS','AVELLANAS CRUDAS','','Origen:Turquía;Categoria:i;Calibre:13/15;Variedad:avellana grano;Puede contener trazas de:CACAHUETES,FRUTOS DE CÁSCARA','','','','',505,0,6,0,0,0,0,'',12.00,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42176.png','','',0,0.000,0),(12308,42178,'NUECES CRUDAS PELADAS','NUECES CRUDAS PELADAS','','Ingredientes:Puede contener trazas de:CACAHUETES,FRUTOS DE CÁSCARA;Origen:USA;Variedad:nuez grano;Categoria:primera','','','','',505,0,7,0,0,0,0,'',15.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42178.png','','',0,0.000,0),(12309,42180,'PISTACHOS TOSTADOS GIGANT','PISTACHOS TOSTADOS GIGANT','','Ingredientes:Puede contener trazas de:CACAHUETES,FRUTOS DE CÁSCARA;Origen:iran;Calibre:20/22;Categoria:primera','','','','',505,0,8,0,0,0,0,'',16.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42180.png','','',0,0.000,0),(12310,42187,'CHULETON VACUNO MADURADO','CHULETON VACUNO MADURADO','','','','','','',201,0,25,1,0,0,0,'',16.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'42187.png','','',0,0.000,0),(12311,42201,'QUESO D.O. CABRALES 1/2','QUESO D.O. CABRALES 1/2','','','','','','',302,0,54,1,0,0,0,'',17.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'42201.png','','',0,0.000,0),(12312,42217,'LECHUGA ICEBERG BOLSA','LECHUGA ICEBERG BOLSA','','','','','','',502,0,119,0,0,0,0,'',0.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42217.png','','',0,0.000,0),(12313,42233,'PAN SALUD SIN SAL','PAN SALUD SIN SAL','','Ingredientes:HARINA DE TRIGO,agua,levadura,emulgentes,E-472e','','','','',601,0,50,0,0,0,0,'',0.45,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'42233.png','','',0,0.000,0),(12314,42234,'PAN GRAN CRISTALINO MASA','PAN GRAN CRISTALINO MASA','','Ingredientes:HARINA DE TRIGO,agua,SEMOLA DE TRIGO,MASA MADRE DE TRIGO,agua,aceite de oliva,sal,levadura,HARINA DE Malta DE TRIGO;Puede contener trazas de:HUEVO,SOJA,LECHE,SESAMO,FRUTOS DE CÁSCARA','','','','',601,0,51,0,0,0,0,'',1.65,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'42234.png','','',0,0.000,0),(12315,42237,'PAN BURGER CRISTALINA MAS','PAN BURGER CRISTALINA MAS','','Ingredientes:HARINA DE TRIGO,agua,MASA MADRE DE TRIGO,aceite de oliva,sal,levadura;Puede contener trazas de:HUEVO,SOJA,LECHE,SESAMO','','','','',601,0,52,0,0,0,0,'',0.45,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'42237.png','','',0,0.000,0),(12316,42247,'LOMO EMPERADOR','LOMO EMPERADOR','','','','','','',701,0,14,1,0,0,0,'',13.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'42247.png','','',0,0.000,0),(12317,42429,'MEDIA BAGUETTE CRISTALINA','MEDIA BAGUETTE CRISTALINA','','Ingredientes:HARINA DE TRIGO,agua,MASA MADRE DE TRIGO,agua,aceite de oliva,sal,levadura,Malta DE TRIGO;Puede contener trazas de:HUEVO,SOJA,LECHE,SESAMO','','','','',601,0,53,0,0,0,0,'',0.50,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'42429.png','','',0,0.000,0),(12318,42443,'CABEZA CERDO IBER. C/PIST','CABEZA CERDO IBER. C/PIST','','','','','','',301,0,70,1,0,0,0,'',12.95,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'42443.png','','',0,0.000,0),(12319,42469,'ENSALADA QUINOA Y BROTES','ENSALADA QUINOA Y BROTES','','','','','','',502,0,120,0,0,0,0,'',3.10,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42469.png','','',0,0.000,0),(12320,42474,'BURGER MEAT POLLO C/ ESPI','BURGER MEAT POLLO C/ ESPI','','','','','','',201,0,26,1,0,0,0,'',7.99,0,0.00,1,0,2,0,'',3,0,0.000,'','-',0,'','-',0,'42474.png','','',0,0.000,0),(12321,42475,'BURGER MEAT MIXTA BARBACO','BURGER MEAT MIXTA BARBACO','','','','','','',201,0,27,1,0,0,0,'',7.99,0,0.00,1,0,2,0,'',3,0,0.000,'','-',0,'','-',0,'42475.png','','',0,0.000,0),(12322,42476,'BURGER MEAT MIXTA JALAPEÑ','BURGER MEAT MIXTA JALAPEÑ','','','','','','',201,0,28,1,0,0,0,'',7.99,0,0.00,1,0,2,0,'',3,0,0.000,'','-',0,'','-',0,'42476.png','','',0,0.000,0),(12323,42492,'TARTA CREMOSA ARANDANOS','TARTA CREMOSA ARANDANOS','','','','','','',603,0,44,0,0,0,0,'',11.50,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42492.png','','',0,0.000,0),(12324,42494,'COULANTS CHOCOLATE','COULANTS CHOCOLATE','','','','','','',603,0,45,0,0,0,0,'',2.70,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42494.png','','',0,0.000,0),(12325,42495,'COULANTS CHOCOLATE FRAMBU','COULANTS CHOCOLATE FRAMBU','','','','','','',603,0,46,0,0,0,0,'',2.70,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42495.png','','',0,0.000,0),(12326,42523,'KALE','KALE','','','','','','',502,0,121,0,0,0,0,'',1.29,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42523.png','','',0,0.000,0),(12327,42524,'TOMATE RAF SELECCION','TOMATE RAF SELECCION','','','','','','',502,0,122,0,0,0,0,'',4.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42524.png','','',0,0.000,0),(12328,42567,'JAMON COCIDO EXTRA EN SU','JAMON COCIDO EXTRA EN SU','','','','','','',301,0,71,1,0,0,0,'',7.95,0,0.00,1,0,3,0,'',15,0,0.000,'','-',0,'','-',0,'42567.png','','',0,0.000,0),(12329,42568,'ENSALADA ENROLLADOS DE PO','ENSALADA ENROLLADOS DE PO','','','','','','',502,0,123,0,0,0,0,'',3.60,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42568.png','','',0,0.000,0),(12330,42569,'TARTA SELVA NEGRA','TARTA SELVA NEGRA','','','','','','',603,0,47,0,0,0,0,'',8.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42569.png','','',0,0.000,0),(12331,42570,'TARTA SAN MARCOS','TARTA SAN MARCOS','','','','','','',603,0,48,0,0,0,0,'',8.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42570.png','','',0,0.000,0),(12332,42571,'BRAZO NATA Y YEMA','BRAZO NATA Y YEMA','','','','','','',603,0,49,0,0,0,0,'',4.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42571.png','','',0,0.000,0),(12333,42572,'BRAZO CHOCOLATE','BRAZO CHOCOLATE','','','','','','',603,0,50,0,0,0,0,'',4.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42572.png','','',0,0.000,0),(12334,42573,'MOUSSE DE LIMON','MOUSSE DE LIMON','','','','','','',603,0,51,0,0,0,0,'',4.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42573.png','','',0,0.000,0),(12335,42574,'MOUSSE DE CHOCOLATE','MOUSSE DE CHOCOLATE','','','','','','',603,0,52,0,0,0,0,'',4.99,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42574.png','','',0,0.000,0),(12336,42575,'TARTA GALLETAS CON CHOCOL','TARTA GALLETAS CON CHOCOL','','','','','','',603,0,53,0,0,0,0,'',8.50,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42575.png','','',0,0.000,0),(12337,42576,'TARTA INFANTIL CHOCOLATE','TARTA INFANTIL CHOCOLATE','','','','','','',603,0,54,0,0,0,0,'',9.30,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42576.png','','',0,0.000,0),(12338,42577,'LIONESAS NATA 10U','LIONESAS NATA 10U','','','','','','',603,0,55,0,0,0,0,'',3.75,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'42577.png','','',0,0.000,0),(12339,42620,'GAMBA COCIDA 300GR','GAMBA COCIDA 300GR','','','','','','',702,0,19,1,0,0,0,'',2.99,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'42620.png','','',0,0.000,0),(12340,42706,'QUESO MEZCLA VIEJO LECHE','QUESO MEZCLA VIEJO LECHE','','','','','','',302,0,55,1,0,0,0,'',12.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'42706.png','','',0,0.000,0),(12341,42713,'CARDOS MANOJO','CARDOS MANOJO','','','','','','',502,0,124,0,0,0,0,'',1.85,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'42713.png','','',0,0.000,0),(12342,42745,'CHIKITO (POLLO FRITO)','CHIKITO (POLLO FRITO)','','','','','','',208,0,23,1,0,0,0,'',3.55,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'42745.png','','',0,0.000,0),(12343,42821,'COSTILLA IBERICA ARGENTIN','COSTILLA IBERICA ARGENTIN','','','','','','',208,0,24,1,0,0,0,'',5.90,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'42821.png','','',0,0.000,0),(12344,42858,'JAMON CEBO IBERICO JABUGO','JAMON CEBO IBERICO JABUGO','','','','','','',301,0,72,1,0,0,0,'',175.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'42858.png','','',0,0.000,0),(12345,42860,'JAMON PAVO CON TRUFA Y BO','JAMON PAVO CON TRUFA Y BO','','','','','','',301,0,73,1,0,0,0,'',8.95,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'42860.png','','',0,0.000,0),(12346,42868,'ALITAS ASADAS PICANTES','ALITAS ASADAS PICANTES','','','','','','',208,0,25,1,0,0,0,'',4.75,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'42868.png','','',0,0.000,0),(12347,43036,'BARRA RUSTICA MASA MADRE','BARRA RUSTICA MASA MADRE','','Ingredientes:HARINA DE TRIGO,agua;Masa madre:HARINA FERMENTADA DE TRIGO,sal,levadura','','','','',601,0,54,0,0,0,0,'',0.60,0,0.00,0,0,6,0,'',1,0,0.000,'','-',0,'','-',0,'43036.png','','',0,0.000,0),(12348,43092,'PRETZEL SALADO','PRETZEL SALADO','','Ingredientes:HARINA DE TRIGO,agua,aceite de nabina,levadura,sal,HARINA Malta CEBADA,GLUTEN DE TRIGO,azucar,estabilizador,E-412,emulgentes,E-472e,acidulante,E-170,E-524','','','','',602,0,49,0,0,0,0,'',0.65,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'43092.png','','',0,0.000,0),(12349,43093,'BRETZEL CREMA ALMENDRA','BRETZEL CREMA ALMENDRA','','Ingredientes:Masa:HARINA DE TRIGO,grasa de palma,aceite de soja y girasol,agua,emulgentes,E-471,acidulante,E-330,E-504,aromas,colorante,E-160,E-171,HUEVO,azucar,levadura,sal;Relleno:pulpa de calabaza,glucosa,espesante,E-407,E-1420,E-1414,conservador,E-223,SULFITOS,E-282,SUERO LACTEO,dextrosa,estabilizante,E-401,E-450,E-339,E-407,aceite de coco,ALMENDRAS,aceite de nabina,PROTEINAS LACTEAS,LACTOSA,almidon,FIBRA DE TRIGO;Contiene trazas de:PESCADO,SOJA,APIO,SESAMO','','','','',602,0,50,0,0,0,0,'',1.00,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'43093.png','','',0,0.000,0),(12350,43094,'BRETZEL CHOCOLATE','BRETZEL CHOCOLATE','','Ingredientes:Masa:HARINA DE TRIGO,grasa de palma,aceite de soja y girasol,agua,emulgentes,E-471,acidulante,E-330,E-263,aromas,colorante,E-160,HUEVO,azucar,levadura,sal;Relleno:crema de chocolate,azucar,cacao,manteca de cacao,espesante,E-1414,glucosa,LECHE EN POLVO,estabilizador,E-450,E-339,cafe,copos de chocolate,emulgentes,E-322 (SOJA),ALMENDRAS,aceite de nabina,PROTEINAS LACTEAS,LACTOSA,almidon,FIBRA DE TRIGO;Contiene trazas de:PESCADO,APIO,SESAMO','','','','',602,0,51,0,0,0,0,'',1.00,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'43094.png','','',0,0.000,0),(12351,43095,'COCA AZUCAR CASERA','COCA AZUCAR CASERA','','','','','','',602,0,52,0,0,0,0,'',1.75,0,0.00,0,0,6,0,'',10,0,0.000,'','-',0,'','-',0,'43095.png','','',0,0.000,0),(12352,43096,'COCA DE NUECES CASERA','COCA DE NUECES CASERA','','','','','','',602,0,53,0,0,0,0,'',2.25,0,0.00,0,0,6,0,'',10,0,0.000,'','-',0,'','-',0,'43096.png','','',0,0.000,0),(12353,43097,'COCA DE ALMENDRA CASERA','COCA DE ALMENDRA CASERA','','','','','','',602,0,54,0,0,0,0,'',2.25,0,0.00,0,0,6,0,'',10,0,0.000,'','-',0,'','-',0,'43097.png','','',0,0.000,0),(12354,43098,'COCA CHOCOLATE CASERA','COCA CHOCOLATE CASERA','','','','','','',602,0,55,0,0,0,0,'',2.25,0,0.00,0,0,6,0,'',10,0,0.000,'','-',0,'','-',0,'43098.png','','',0,0.000,0),(12355,43099,'COCA MANZANA CASERA','COCA MANZANA CASERA','','','','','','',602,0,56,0,0,0,0,'',2.25,0,0.00,0,0,6,0,'',10,0,0.000,'','-',0,'','-',0,'43099.png','','',0,0.000,0),(12356,43100,'COCA CALABAZA CASERA','COCA CALABAZA CASERA','','','','','','',602,0,57,0,0,0,0,'',2.25,0,0.00,0,0,6,0,'',6,0,0.000,'','-',0,'','-',0,'43100.png','','',0,0.000,0),(12357,43114,'PINCHOS MORUNOS POLLO','PINCHOS MORUNOS POLLO','','','','','','',208,0,26,1,0,0,0,'',9.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43114.png','','',0,0.000,0),(12358,43132,'BURGER MEAT VACUNO','BURGER MEAT VACUNO','','','','','','',201,0,29,1,0,0,0,'',9.50,0,0.00,1,0,2,0,'',3,0,0.000,'','-',0,'','-',0,'43132.png','','',0,0.000,0),(12359,43143,'ENSALADA GOURMET','ENSALADA GOURMET','','','','','','',502,0,125,0,0,0,0,'',2.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43143.png','','',0,0.000,0),(12360,43148,'CENTRO BABILLA WAGYU SKIN','CENTRO BABILLA WAGYU SKIN','','','','','','',201,0,30,1,0,0,0,'',10.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43148.png','','',0,0.000,0),(12361,43176,'JAMON BODEGA SELECCION PI','JAMON BODEGA SELECCION PI','','','','','','',301,0,74,1,0,0,0,'',42.00,0,0.00,0,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'43176.png','','',0,0.000,0),(12362,43329,'NUECES PIZARRO','NUECES PIZARRO','','Variedad:pizarro;Origen:España;Calibre:32+;Categoria:i','','','','',505,0,4,1,0,0,0,'',5.99,0,0.00,1,63,4,0,'',0,0,0.000,'','-',0,'','-',0,'43329.png','','',0,0.000,0),(12363,43377,'TRONCO CHOCOLATE','TRONCO CHOCOLATE','','','','','','',603,0,56,0,0,0,0,'',2.25,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'43377.png','','',0,0.000,0),(12364,43381,'MILHOJAS CERDO IBERICO','MILHOJAS CERDO IBERICO','','','','','','',208,0,27,1,0,0,0,'',12.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43381.png','','',0,0.000,0),(12365,43394,'COCTEL FRUTOS SECOS NATUR','COCTEL FRUTOS SECOS NATUR','','Puede contener trazas de:CACAHUETES,FRUTOS DE CÁSCARA','','','','',505,0,10,0,0,0,0,'',9.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43394.png','','',0,0.000,0),(12366,43395,'ALMENDRAS TOSTADAS CASCAR','ALMENDRAS TOSTADAS CASCAR','','Ingredientes:Puede contener trazas de:CACAHUETES,FRUTOS DE CÁSCARA;Origen:españa/Chile/usa;Variedad:mollar/non pareil,','','','','',505,0,11,0,0,0,0,'',12.00,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43395.png','','',0,0.000,0),(12367,43397,'COCTEL FRUTOS SECOS ENSAL','COCTEL FRUTOS SECOS ENSAL','','Ingredientes:Puede contener trazas de:CACAHUETES,FRUTOS DE CÁSCARA;Contiene:SULFITOS','','','','',505,0,12,0,0,0,0,'',9.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43397.png','','',0,0.000,0),(12368,43430,'JAMON WESTFALIA ASADO','JAMON WESTFALIA ASADO','','','','','','',301,0,75,1,0,0,0,'',9.95,0,0.00,1,0,3,0,'',15,0,0.000,'','-',0,'','-',0,'43430.png','','',0,0.000,0),(12369,43459,'ASADO POLLO CON MOUSSE DE','ASADO POLLO CON MOUSSE DE','','','','','','',301,0,76,1,0,0,0,'',15.95,0,0.00,1,0,3,0,'',30,0,0.000,'','-',0,'','-',0,'43459.png','','',0,0.000,0),(12370,43469,'POLLO PAZO DE PIAS','POLLO PAZO DE PIAS','','','','','','',205,0,31,1,0,0,0,'',3.75,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43469.png','','',0,0.000,0),(12371,43487,'CANONIGOS','CANONIGOS','','','','','','',502,0,126,0,0,0,0,'',1.00,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43487.png','','',0,0.000,0),(12372,43505,'PATE DE OPORTO','PATE DE OPORTO','','','','','','',301,0,77,1,0,0,0,'',12.96,0,0.00,1,0,3,0,'',10,0,0.000,'','-',0,'','-',0,'43505.png','','',0,0.000,0),(12373,43722,'ENSALADA ROMA','ENSALADA ROMA','','','','','','',502,0,127,0,0,0,0,'',2.59,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43722.png','','',0,0.000,0),(12374,43723,'CEBOLLA ECHALOTE','CEBOLLA ECHALOTE','','','','','','',502,0,128,0,0,0,0,'',1.29,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43723.png','','',0,0.000,0),(12375,43724,'PISTOLA DESHUESADA GALLEG','PISTOLA DESHUESADA GALLEG','','','','','','',201,0,31,1,0,0,0,'',12.80,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43724.png','','',0,0.000,0),(12376,43725,'CADERA-BABILLA GALLEGA','CADERA-BABILLA GALLEGA','','','','','','',201,0,32,1,0,0,0,'',17.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43725.png','','',0,0.000,0),(12377,43726,'FILETE TIERNO GALLEGA','FILETE TIERNO GALLEGA','','','','','','',201,0,33,1,0,0,0,'',14.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43726.png','','',0,0.000,0),(12378,43727,'ENTRECOT TERNERA GALLEGA','ENTRECOT TERNERA GALLEGA','','','','','','',201,0,34,1,0,0,0,'',22.98,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43727.png','','',0,0.000,0),(12379,43728,'CHULETON TERNERA GALLEGA','CHULETON TERNERA GALLEGA','','','','','','',201,0,35,1,0,0,0,'',19.95,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43728.png','','',0,0.000,0),(12380,43729,'SOLOMILLO TERNERA GALLEGA','SOLOMILLO TERNERA GALLEGA','','','','','','',201,0,36,1,0,0,0,'',33.70,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43729.png','','',0,0.000,0),(12381,43733,'TERNERA GUISAR GALLEGA','TERNERA GUISAR GALLEGA','','','','','','',201,0,37,1,0,0,0,'',10.95,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43733.png','','',0,0.000,0),(12382,43745,'AGUACATE ECOLOGICO','AGUACATE ECOLOGICO','','','','','','',501,0,55,0,0,0,0,'',2.90,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43745.png','','',0,0.000,0),(12383,43746,'ENSALADA DEL HUERTO','ENSALADA DEL HUERTO','','','','','','',502,0,129,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'43746.png','','',0,0.000,0),(12384,43747,'BAGUETTE 100% INTEGRAL','BAGUETTE 100% INTEGRAL','','Ingredientes:HARINA INTEGRAL DE TRIGO,agua,masa madre de trigo,sal,harina integral de cebada,GLUTEN,levadura;Puede contener trazas de:HUEVO,SOJA,LECHE,FRUTOS DE CÁSCARA,SESAMO','','','','',601,0,55,0,0,0,0,'',1.69,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'43747.png','','',0,0.000,0),(12385,43748,'BOCATA PIAMONTE','BOCATA PIAMONTE','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura;Emulgentes:E-472e;Puede contener trazas de:SESAMO,ALTRAMUCES','','','','',601,0,56,0,0,0,0,'',0.27,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'43748.png','','',0,0.000,0),(12386,43750,'MINI MALLAS SALADAS','MINI MALLAS SALADAS','','Ingredientes:Masa:HARINA DE TRIGO;Margarina:aceite de palma,aceite de soja,aceite de girasol,jarabe de glucosa-fructosa;Emulgentes:E-471,E-322;Colorantes:E-160a,agua,levadura,HUEVO,azucar,sal,LECHE EN POLVO,GLUTEN;Emulgentes:E-471,E-472e,E-322,grasa de palma,dextrosa,HARINA DE TRIGO;Agente de tratamiento de la harina:E-300;Relleno:espinacas,aceite de oliva,aceite de girasol,sal,aromas,almidon de maiz,CLARA DE HUEVO,pasta de sobrasada,queso emental,crema de queso,fiambre de cerdo;Contiene:SOJA Y DERIVADOS;Puede contener trazas de:FRUTOS DE CÁSCARA','','','','',602,0,58,0,0,0,0,'',0.24,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'43750.png','','',0,0.000,0),(12387,43751,'MINI NORDICO AVENA','MINI NORDICO AVENA','','Ingredientes:HARINA DE TRIGO,agua,sal,levadura;Emulgentes:E-472e;Cobertura:copos de avena,pipas de girasol,lino dorado,lino marron,SESAMO,semillas de amapola;Puede contener trazas de:SOJA,ALTRAMUCES','','','','',601,0,57,0,0,0,0,'',0.22,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'43751.png','','',0,0.000,0),(12388,43752,'TOMATE MORESCO','TOMATE MORESCO','','Origen:España;Variedad:moresco;Calibre:m/g;Categoria:i','','','','',502,0,14,1,0,0,0,'',2.99,0,0.00,1,57,4,0,'',0,0,0.000,'','-',0,'','-',0,'43752.png','','',0,0.000,0),(12389,43780,'SOBRASADA IBERICA SELECTA','SOBRASADA IBERICA SELECTA','','','','','','',301,0,78,1,0,0,0,'',7.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'43780.png','','',0,0.000,0),(12390,43800,'HOGAZA REBANADA NORMAL','HOGAZA REBANADA NORMAL','','','','','','',601,0,58,0,0,0,0,'',1.20,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'43800.png','','',0,0.000,0),(12391,43801,'HOGAZA 100% INTEGRAL REBA','HOGAZA 100% INTEGRAL REBA','','','','','','',601,0,59,0,0,0,0,'',1.40,0,0.00,0,0,6,0,'',0,0,0.000,'','-',0,'','-',0,'43801.png','','',0,0.000,0),(12392,43813,'CHULETAS SAJONIA','CHULETAS SAJONIA','','','','','','',208,0,28,1,0,0,0,'',9.75,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43813.png','','',0,0.000,0),(12393,43897,'LUBINA UNIDAD','LUBINA UNIDAD','','','','','','',701,0,15,1,0,0,0,'',4.75,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'43897.png','','',0,0.000,0),(12394,43898,'MEJILLON MALLA 1KG','MEJILLON MALLA 1KG','','','','','','',702,0,20,1,0,0,0,'',2.40,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'43898.png','','',0,0.000,0),(12395,43900,'FRUTOS SECOS GRANEL','FRUTOS SECOS GRANEL','','','','','','',505,0,2,1,0,0,0,'',16.50,0,0.00,1,20,4,0,'',0,0,0.000,'','-',0,'','-',0,'43900.png','','',0,0.000,0),(12396,43912,'BURGER POLLO ESPINACA NAT','BURGER POLLO ESPINACA NAT','','','','','','',208,0,29,1,0,0,0,'',2.89,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'43912.png','','',0,0.000,0),(12397,43914,'TORTA NARANJA CHOCOLATE','TORTA NARANJA CHOCOLATE','','','','','','',602,0,59,0,0,0,0,'',3.39,0,0.00,0,0,6,0,'',7,0,0.000,'','-',0,'','-',0,'43914.png','','',0,0.000,0),(12398,43915,'COCA ESCUDELLÀ CALABAZA','COCA ESCUDELLÀ CALABAZA','','','','','','',602,0,60,0,0,0,0,'',2.36,0,0.00,0,0,6,0,'',7,0,0.000,'','-',0,'','-',0,'43915.png','','',0,0.000,0),(12399,43916,'COCA ESCUDELLÀ ALMENDRA','COCA ESCUDELLÀ ALMENDRA','','','','','','',602,0,61,0,0,0,0,'',2.98,0,0.00,0,0,6,0,'',7,0,0.000,'','-',0,'','-',0,'43916.png','','',0,0.000,0),(12400,43920,'ALMEJA JAPONICA MALLA 1/2','ALMEJA JAPONICA MALLA 1/2','','','','','','',702,0,21,1,0,0,0,'',12.00,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'43920.png','','',0,0.000,0),(12401,44058,'POLLO ASADO','POLLO ASADO','','','','','','',208,0,30,1,0,0,0,'',5.30,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44058.png','','',0,0.000,0),(12402,44060,'PUNTILLA CONGELADA','PUNTILLA CONGELADA','','','','','','',702,0,22,1,0,0,0,'',5.00,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44060.png','','',0,0.000,0),(12403,44065,'QUESO OVEJA AÑEJO PIMIENT','QUESO OVEJA AÑEJO PIMIENT','','','','','','',302,0,56,1,0,0,0,'',17.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'44065.png','','',0,0.000,0),(12404,44066,'QUESO OVEJA AÑEJO TEMPRAN','QUESO OVEJA AÑEJO TEMPRAN','','','','','','',302,0,57,1,0,0,0,'',17.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'44066.png','','',0,0.000,0),(12405,44067,'QUESO CURADO MAESTRIA','QUESO CURADO MAESTRIA','','','','','','',302,0,58,1,0,0,0,'',12.99,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'44067.png','','',0,0.000,0),(12406,44071,'ENSALADA CESAR LIGHT','ENSALADA CESAR LIGHT','','','','','','',502,0,131,0,0,0,0,'',2.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'44071.png','','',0,0.000,0),(12407,44102,'RODABALLO FRESCO','RODABALLO FRESCO','','','','','','',701,0,16,1,0,0,0,'',15.00,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44102.png','','',0,0.000,0),(12408,44175,'GALERAS CONGELADAS','GALERAS CONGELADAS','','','','','','',702,0,23,1,0,0,0,'',8.80,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44175.png','','',0,0.000,0),(12409,44177,'GAMBA ALISTADA CRUDA 300G','GAMBA ALISTADA CRUDA 300G','','','','','','',702,0,24,1,0,0,0,'',3.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44177.png','','',0,0.000,0),(12410,44178,'GAMBA BLANCA CRUDA 300GR','GAMBA BLANCA CRUDA 300GR','','','','','','',702,0,25,1,0,0,0,'',2.99,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44178.png','','',0,0.000,0),(12411,44180,'PAVO SANDWICH BARRA','PAVO SANDWICH BARRA','','','','','','',301,0,79,1,0,0,0,'',5.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'44180.png','','',0,0.000,0),(12412,44190,'MEDIO POLLO ASADO','MEDIO POLLO ASADO','','','','','','',2004,0,1,1,0,0,0,'',4.00,0,0.00,0,11,8,0,'',0,0,0.000,'','-',0,'','-',0,'44190.png','','',0,0.000,0),(12413,44191,'POLLO ASADO','POLLO ASADO','','','','','','',2004,0,2,1,0,0,0,'',7.50,0,0.00,0,12,8,0,'',0,0,0.000,'','-',0,'','-',0,'44191.png','','',0,0.000,0),(12414,44192,'RACION ARROZ DEL DIA','RACION ARROZ DEL DIA','','','','','','',2001,0,1,1,0,0,0,'',3.00,0,0.00,0,13,8,0,'',0,0,0.000,'','-',0,'','-',0,'44192.png','','',0,0.000,0),(12415,44193,'RACION CARNE DEL DIA','RACION CARNE DEL DIA','','','','','','',2004,0,3,1,0,0,0,'',3.50,0,0.00,0,14,8,0,'',0,0,0.000,'','-',0,'','-',0,'44193.png','','',0,0.000,0),(12416,44194,'RACION ENSALADILLA DEL DI','RACION ENSALADILLA DEL DI','','','','','','',2008,0,1,1,0,0,0,'',2.50,0,0.00,0,15,8,0,'',0,0,0.000,'','-',0,'','-',0,'44194.png','','',0,0.000,0),(12417,44195,'RACION PASTA DEL DIA','RACION PASTA DEL DIA','','','','','','',2002,0,1,1,0,0,0,'',3.00,0,0.00,0,16,8,0,'',0,0,0.000,'','-',0,'','-',0,'44195.png','','',0,0.000,0),(12418,44196,'RACION PESCADO DEL DIA','RACION PESCADO DEL DIA','','','','','','',2005,0,1,1,0,0,0,'',4.00,0,0.00,0,17,8,0,'',0,0,0.000,'','-',0,'','-',0,'44196.png','','',0,0.000,0),(12419,44197,'RACION CREMA DEL DIA','RACION CREMA DEL DIA','','','','','','',2003,0,1,1,0,0,0,'',2.00,0,0.00,0,18,8,0,'',0,0,0.000,'','-',0,'','-',0,'44197.png','','',0,0.000,0),(12420,44207,'HAMBURGUESA PULPO Y CORVI','HAMBURGUESA PULPO Y CORVI','','','','','','',701,0,17,1,0,0,0,'',3.95,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44207.png','','',0,0.000,0),(12421,44208,'HAMBURGUESA MEDITERRANEA','HAMBURGUESA MEDITERRANEA','','','','','','',701,0,18,1,0,0,0,'',3.95,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44208.png','','',0,0.000,0),(12422,44209,'LONGANIZA MARINA','LONGANIZA MARINA','','','','','','',701,0,19,1,0,0,0,'',5.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44209.png','','',0,0.000,0),(12423,44293,'JAMONCITOS ASADOS','JAMONCITOS ASADOS','','','','','','',208,0,31,1,0,0,0,'',3.15,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44293.png','','',0,0.000,0),(12424,44294,'CUARTOS TRASEROS ASADOS','CUARTOS TRASEROS ASADOS','','','','','','',208,0,32,1,0,0,0,'',3.15,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44294.png','','',0,0.000,0),(12425,44295,'ALAS ASADAS','ALAS ASADAS','','','','','','',208,0,33,1,0,0,0,'',2.49,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44295.png','','',0,0.000,0),(12426,44320,'ALBONDIGAS DE POLLO','ALBONDIGAS DE POLLO','','','','','','',208,0,34,1,0,0,0,'',2.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44320.png','','',0,0.000,0),(12427,44321,'PECHUGA VILLAROY','PECHUGA VILLAROY','','','','','','',208,0,35,1,0,0,0,'',3.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44321.png','','',0,0.000,0),(12428,44322,'BURGUER POLLO-QUESO','BURGUER POLLO-QUESO','','','','','','',208,0,36,1,0,0,0,'',3.48,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44322.png','','',0,0.000,0),(12429,44433,'DOTS DUO P-4','DOTS DUO P-4','','Contiene:GLUTEN,SOJA Y DERIVADOS,LECHE Y DERIVADOS;Puede contener trazas de:HUEVO Y DERIVADOS,FRUTOS DE CÁSCARA','','','','',602,0,62,0,0,0,0,'',1.59,0,0.00,0,0,6,0,'',3,0,0.000,'','-',0,'','-',0,'44433.png','','',0,0.000,0),(12430,44485,'CARACOL CABRILLA MALLA 1/','CARACOL CABRILLA MALLA 1/','','','','','','',702,0,26,1,0,0,0,'',3.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44485.png','','',0,0.000,0),(12431,44525,'CHULETA LOMO CON PIEL','CHULETA LOMO CON PIEL','','','','','','',202,0,30,1,0,0,0,'',6.25,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44525.png','','',0,0.000,0),(12432,44610,'FILETE LUBINA UNIDAD','FILETE LUBINA UNIDAD','','','','','','',701,0,20,1,0,0,0,'',3.60,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44610.png','','',0,0.000,0),(12433,44617,'EMPANADILLA TOMATE','EMPANADILLA TOMATE','','Ingredientes:Relleno:tomate,pimiento verde,pimiento rojo,azucar,sal,aceite de oliva,HUEVO,CABALLA,PAN RALLADO;Masa:HARINA DE TRIGO,agua con gas,aceite de girasol,mejorante,enzimas,E-170 (i),vinagre;Puede contener trazas de:LECHE,FRUTOS DE CÁSCARA,SESAMO','','','','',602,0,63,0,0,0,0,'',0.90,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'44617.png','','',0,0.000,0),(12434,44618,'MINI EMPANADILLA TOMATE','MINI EMPANADILLA TOMATE','','Ingredientes:Relleno:tomate,pimiento verde,pimiento rojo,azucar,sal,aceite de oliva,HUEVO,CABALLA,PAN RALLADO;Masa:HARINA DE TRIGO,agua con gas,aceite de girasol,mejorante,enzimas,E-170 (i),vinagre;Puede contener trazas de:LECHE,FRUTOS DE CÁSCARA,SESAMO','','','','',602,0,64,0,0,0,0,'',0.33,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'44618.png','','',0,0.000,0),(12435,44656,'MINI EMPANADILLA GUISANTE','MINI EMPANADILLA GUISANTE','','Ingredientes:HARINA DE TRIGO;Relleno:cebolla,guisantes,ATUN,HUEVO,aceite de oliva,agua,sal,agua con gas,aceite de girasol,mejorante,enzimas,,vinagre;Puede contener trazas de:LECHE,FRUTOS DE CÁSCARA,SESAMO','','','','',602,0,65,0,0,0,0,'',0.33,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'44656.png','','',0,0.000,0),(12436,44657,'MINI EMPANADILLA ESPINACA','MINI EMPANADILLA ESPINACA','','Ingredientes:HARINA DE TRIGO;Relleno:espinacas,aceite de girasol,ajo,sal,almidon modificado,E-1422,regulador de acidez,acido citrico,conservador,E-202,agua,harina de arroz,LECITINA DE SOJA,pimenton,mejorante,agente tratamiento harina;Contiene:HUEVO;Puede contener trazas de:PESCADO,FRUTOS DE CÁSCARA,SESAMO,LECHE','','','','',602,0,66,0,0,0,0,'',0.33,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'44657.png','','',0,0.000,0),(12437,44689,'ZAMBURIÑAS','ZAMBURIÑAS','','','','','','',702,0,27,1,0,0,0,'',22.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44689.png','','',0,0.000,0),(12438,44690,'ZAMBURIÑA UNIDAD','ZAMBURIÑA UNIDAD','','','','','','',702,0,28,1,0,0,0,'',0.99,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44690.png','','',0,0.000,0),(12439,44698,'LANGOSTINO RECIEN COCIDO','LANGOSTINO RECIEN COCIDO','','','','','','',702,0,29,1,0,0,0,'',14.95,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44698.png','','',0,0.000,0),(12440,44718,'PECHUGA DE POLLO COCIDA','PECHUGA DE POLLO COCIDA','','','','','','',301,0,80,1,0,0,0,'',10.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'44718.png','','',0,0.000,0),(12441,44747,'COCA TOMATE','COCA TOMATE','','Ingredientes:HARINA DE TRIGO,agua,tomate,cebolla,pimiento,calabacin,berenjena,ajo,aceite de oliva,aceite de girasol,levadura,sal,azucar,antiapelmazante,E-170,emulgentes,E-472e,E-322,acidulante,E-296,E-334,regulador de acidez,E-330,E-331,E-331(iii),E-332,antioxidante,E-300,almidon modificado,E-1422,espesante,E-461,E-410,bicarbonato sodico,enzimas;Puede contener trazas de:HUEVO,PESCADO,SOJA,LECHE,FRUTOS DE CÁSCARA,SESAMO,SULFITOS,ALTRAMUCES','','','','',602,0,67,0,0,0,0,'',0.90,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'44747.png','','',0,0.000,0),(12442,44748,'COCA VERDURA INTEGRAL','COCA VERDURA INTEGRAL','','Ingredientes:HARINA DE TRIGO,HARINA INTEGRAL DE TRIGO,agua,cebolla,tomate,pimiento,aceite de oliva,aceite de girasol,levadura,sal,azucar,antiapelmazante,E-170,emulgentes,E-472e,E-322,acidulante,E-296,E-334,regulador de acidez,E-330,E-331,E-331(iii),E-332,bicarbonato sodico,antioxidante,E-300,enzimas,almidon modificado,E-1422,espesante,E-410,conservantes,E-202;Puede contener trazas de:HUEVO,PESCADO,SOJA,LECHE,FRUTOS DE CÁSCARA,SESAMO,SULFITOS,ALTRAMUCES','','','','',602,0,68,0,0,0,0,'',0.90,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'44748.png','','',0,0.000,0),(12443,44749,'COCA GUISANTES','COCA GUISANTES','','Ingredientes:HARINA DE TRIGO,agua,guisantes,cebolla,tomate,aceite de oliva,aceite de girasol,levadura,sal,azucar,antiapelmazante,E-170,emulgentes,E-472e,E-322,antioxidante,E-300,acidulante,E-296,E-334,regulador de acidez,E-330,E-331,E-331(iii),E-332,conservantes,E-202,bicarbonato sodico,enzimas,almidon modificado,E-1422,espesante,E-410;Puede contener trazas de:HUEVO,PESCADO,SOJA,LECHE,FRUTOS DE CÁSCARA,SESAMO,SULFITOS,ALTRAMUCES','','','','',602,0,69,0,0,0,0,'',0.85,0,0.00,0,0,6,0,'',2,0,0.000,'','-',0,'','-',0,'44749.png','','',0,0.000,0),(12444,44849,'ESPALDILLA VACUNO SKIN','ESPALDILLA VACUNO SKIN','','','','','','',201,0,38,1,0,0,0,'',13.90,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44849.png','','',0,0.000,0),(12445,44850,'DELICIAS SOLOMILLO POLLO','DELICIAS SOLOMILLO POLLO','','','','','','',208,0,37,1,0,0,0,'',1.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44850.png','','',0,0.000,0),(12446,44852,'NUGGETS DE POLLO','NUGGETS DE POLLO','','','','','','',208,0,38,1,0,0,0,'',2.50,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44852.png','','',0,0.000,0),(12447,44854,'CORDON BLEU POLLO','CORDON BLEU POLLO','','','','','','',208,0,39,1,0,0,0,'',2.66,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44854.png','','',0,0.000,0),(12448,44855,'FINGERS PECHUGA POLLO','FINGERS PECHUGA POLLO','','','','','','',208,0,40,1,0,0,0,'',2.49,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'44855.png','','',0,0.000,0),(12449,44894,'GUACAMOLE','GUACAMOLE','','','','','','',502,0,132,0,0,0,0,'',1.85,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'44894.png','','',0,0.000,0),(12450,44918,'LANGOSTINO COCIDO 80/100','LANGOSTINO COCIDO 80/100','','','','','','',702,0,30,1,0,0,0,'',6.95,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'44918.png','','',0,0.000,0),(12451,45092,'CIGALA BANDEJA 1KG','CIGALA BANDEJA 1KG','','','','','','',702,0,31,1,0,0,0,'',10.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'45092.png','','',0,0.000,0),(12452,45214,'DATILES RAMA','DATILES RAMA','','','','','','',505,0,14,0,0,0,0,'',2.79,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45214.png','','',0,0.000,0),(12453,45231,'DORADA GRANDE UNIDAD','DORADA GRANDE UNIDAD','','','','','','',701,0,21,1,0,0,0,'',4.80,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'45231.png','','',0,0.000,0),(12454,45232,'LANGOSTINO CRUDO 50-60','LANGOSTINO CRUDO 50-60','','','','','','',702,0,32,1,0,0,0,'',9.95,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'45232.png','','',0,0.000,0),(12455,45234,'PECHUGA DE PAVO','PECHUGA DE PAVO','','','','','','',301,0,81,1,0,0,0,'',12.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'45234.png','','',0,0.000,0),(12456,45253,'ENSALADA COMPLETA PASTA I','ENSALADA COMPLETA PASTA I','','','','','','',502,0,133,0,0,0,0,'',2.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45253.png','','',0,0.000,0),(12457,45256,'VIEIRA MEDIA CONCHA','VIEIRA MEDIA CONCHA','','','','','','',702,0,33,1,0,0,0,'',2.45,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'45256.png','','',0,0.000,0),(12458,45257,'PASTISSET CABELLO DE ANGE','PASTISSET CABELLO DE ANGE','','Ingredientes:Cabello de angel:calabaza,azucar,almidon modificado,E-1422,espesante,E-410,E-412,conservador,E-202,E-223,acido citrico,aroma de anis,HARINA DE TRIGO,manteca de cerdo,agua,EXTRACTO DE Malta,aroma;Puede contener trazas de:PESCADO,LECHE,FRUTOS DE CÁSCARA,SOJA,HUEVO','','','','',602,0,70,0,0,0,0,'',8.45,0,0.00,1,0,6,0,'',7,0,0.000,'','-',0,'','-',0,'45257.png','','',0,0.000,0),(12459,45306,'QUESO DE OVEJA CON TRUFA','QUESO DE OVEJA CON TRUFA','','','','','','',302,0,59,1,0,0,0,'',16.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'45306.png','','',0,0.000,0),(12460,45321,'CALABAZA DADOS','CALABAZA DADOS','','','','','','',502,0,134,0,0,0,0,'',1.95,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45321.png','','',0,0.000,0),(12461,45324,'SHIMEJI MARRON','SHIMEJI MARRON','','','','','','',502,0,135,0,0,0,0,'',1.69,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45324.png','','',0,0.000,0),(12462,45325,'SHIMEJI BLANCO','SHIMEJI BLANCO','','','','','','',502,0,136,0,0,0,0,'',1.69,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45325.png','','',0,0.000,0),(12463,45361,'CALABAZA ECOLOGICA','CALABAZA ECOLOGICA','','','','','','',502,0,13,1,0,0,0,'',1.35,0,0.00,1,53,4,0,'',0,0,0.000,'','-',0,'','-',0,'45361.png','','',0,0.000,0),(12464,45362,'PIMIENTO PICANTE','PIMIENTO PICANTE','','','','','','',502,0,138,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45362.png','','',0,0.000,0),(12465,45364,'JAMON CEBO IBERICO DESHUE','JAMON CEBO IBERICO DESHUE','','','','','','',301,0,82,1,0,0,0,'',49.00,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'45364.png','','',0,0.000,0),(12466,45365,'QUESO DE OVEJA MEDITERRAN','QUESO DE OVEJA MEDITERRAN','','','','','','',302,0,60,1,0,0,0,'',13.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'45365.png','','',0,0.000,0),(12467,45371,'TOMATE NATURAL RALLADO EC','TOMATE NATURAL RALLADO EC','','','','','','',502,0,139,0,0,0,0,'',1.69,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45371.png','','',0,0.000,0),(12468,45372,'TOMATE NATURAL RALLADO','TOMATE NATURAL RALLADO','','','','','','',502,0,140,0,0,0,0,'',0.95,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45372.png','','',0,0.000,0),(12469,45373,'PIÑA PELADA RODAJAS','PIÑA PELADA RODAJAS','','','','','','',501,0,56,0,0,0,0,'',2.69,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45373.png','','',0,0.000,0),(12470,45374,'TACOS ATUN CONGELADO','TACOS ATUN CONGELADO','','','','','','',701,0,22,1,0,0,0,'',8.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'45374.png','','',0,0.000,0),(12471,45381,'JAMON SERR. G. RVA MEDIA','JAMON SERR. G. RVA MEDIA','','','','','','',301,0,83,1,0,0,0,'',9.95,0,0.00,1,0,3,0,'',0,0,0.000,'','-',0,'','-',0,'45381.png','','',0,0.000,0),(12472,45382,'PIMIENTO ASADO TIRAS ECOL','PIMIENTO ASADO TIRAS ECOL','','','','','','',502,0,141,0,0,0,0,'',2.09,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45382.png','','',0,0.000,0),(12473,45408,'GRAN BURGUER','GRAN BURGUER','','','','','','',208,0,41,1,0,0,0,'',2.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'45408.png','','',0,0.000,0),(12474,45409,'NUGGETS JAMON Y POLLO','NUGGETS JAMON Y POLLO','','','','','','',208,0,42,1,0,0,0,'',2.99,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'45409.png','','',0,0.000,0),(12475,45445,'MANZANA ROYAL GALA ECOLOG','MANZANA ROYAL GALA ECOLOG','','','','','','',501,0,57,0,0,0,0,'',2.89,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45445.png','','',0,0.000,0),(12476,45446,'MANZANA GOLDEN ECOLOGICA','MANZANA GOLDEN ECOLOGICA','','','','','','',501,0,58,0,0,0,0,'',2.89,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45446.png','','',0,0.000,0),(12477,45493,'PATAS PULPO COCIDAS 2U.','PATAS PULPO COCIDAS 2U.','','','','','','',702,0,34,1,0,0,0,'',12.50,0,0.00,1,0,5,0,'',0,0,0.000,'','-',0,'','-',0,'45493.png','','',0,0.000,0),(12478,45495,'PATATA VIOLETA LAVADA','PATATA VIOLETA LAVADA','','','','','','',502,0,142,0,0,0,0,'',1.99,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45495.png','','',0,0.000,0),(12479,45505,'BURGUER POTRO','BURGUER POTRO','','','','','','',204,0,2,1,0,0,0,'',2.55,0,0.00,1,0,2,0,'',0,0,0.000,'','-',0,'','-',0,'45505.png','','',0,0.000,0),(12480,45515,'ENSALADA CLASICA','ENSALADA CLASICA','','','','','','',502,0,143,0,0,0,0,'',1.49,0,0.00,1,0,4,0,'',0,0,0.000,'','-',0,'','-',0,'45515.png','','',0,0.000,0),(12481,45548,'TOMATE ROSA','TOMATE ROSA','','','','','','',502,0,25,1,0,0,0,'',3.99,0,0.00,1,94,4,0,'',0,0,0.000,'','-',0,'','-',0,'45548.png','','',0,0.000,0),(12482,45549,'PIÑA AVION','PIÑA AVION','','','','','','',501,0,13,1,0,0,0,'',3.49,0,0.00,1,46,4,0,'',0,0,0.000,'','-',0,'','-',0,'45549.png','','',0,0.000,0);
/*!40000 ALTER TABLE `items` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemsInsertedB` BEFORE INSERT ON `Items`
FOR EACH ROW BEGIN
  IF(NEW.`FamilyOrder`=0) THEN
    SET NEW.`FamilyOrder`=IFNULL(
      (SELECT t1.`FamilyOrder`+1 as _id FROM (SELECT `FamilyOrder` FROM `Items` WHERE `Family`=NEW.`Family`) AS t1 
      LEFT JOIN (SELECT `FamilyOrder` FROM `Items` WHERE `Family`=NEW.`Family`) AS t2 ON t2.`FamilyOrder`=t1.`FamilyOrder`+1 
      WHERE t2.`FamilyOrder` IS NULL LIMIT 1), 1);
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemsInserted` AFTER INSERT ON `Items`
FOR EACH ROW BEGIN
  IF(NEW.`EAN13`!='') THEN
    INSERT INTO ItemsEAN13 (`Code`, `EAN13`) VALUES (NEW.`Code`, NEW.`EAN13`) ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
  END IF;
  SET @loadfunc := IFNULL (GLOBAL_GET("PLUload-func"),99);
  CASE @loadfunc

    WHEN 0 THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,`Code` FROM `PLUKeys`
        WHERE `Code`=NEW.`Code` AND (`PLUKeys`.`Section`=GLOBAL_GET("current_group") OR (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard")) )
        ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);

    WHEN 1 THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,`Code` FROM `PLUKeys`
        WHERE `Code`=NEW.`Code` AND (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard"))
        ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,NEW.`Family`,`PLUKey`,`Code` FROM `PLUKeys`
        WHERE `Code`=NEW.`Code` AND `PLUKeys`.`Section`=GLOBAL_GET("current_group")
        ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);

    WHEN 2 THEN
      SET @famor := (SELECT `FamilyIndex` FROM `CounterFamilies` WHERE `CounterCode`=GLOBAL_GET("current_counter") AND `FamilyCode`=NEW.`Family`);
      IF (NEW.`OnKey`=TRUE) THEN
        IF (@famor IS NOT NULL) THEN
          INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES(GLOBAL_GET("current_group"),@famor,NEW.`FamilyOrder`,NEW.`Code`)
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
        END IF;
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,`Code` FROM `PLUKeys`
          WHERE `Code`=NEW.`Code` AND (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard"))
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;

    WHEN 3 THEN
      IF (NEW.`Counter`!=0) AND (NEW.`PLU`!=0) AND (NEW.`Counter`=GLOBAL_GET("current_counter")) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (NEW.`Counter`,0,NEW.`PLU`, NEW.`Code`)
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;

    WHEN 100 THEN
      IF (NEW.`Counter`!=0) AND (NEW.`PLU`!=0) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Counter`,board_idx,NEW.`PLU`, NEW.`Code`
        FROM (SELECT IF(Code=GLOBAL_GET("current_counter"),1,@a:=@a+1) AS board_idx, Code FROM (SELECT @a:=1) AS ct, Counters) AS tt100 WHERE NEW.`Counter`=tt100.Code
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;

    WHEN 101 THEN
      IF (NEW.`Counter`!=0) AND (NEW.`PLU`!=0) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Counter`,board_idx,NEW.`PLU`, NEW.`Code`
        FROM (SELECT @a:=@a+1 AS board_idx, Code FROM (SELECT @a:=0) AS ct, Counters) AS tt100 WHERE NEW.`Counter`=tt100.Code
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;

    WHEN 4 THEN
      IF ((NEW.`PLU`!=0) AND (NEW.`Code`!=0)) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),`ItemGroupKeys`.`NumKey`,NEW.`PLU`,NEW.`Code` FROM `ItemGroupKeys`
          WHERE ((`ItemGroupKeys`.`Abovesection`!=0 AND `ItemGroupKeys`.`Abovesection`=NEW.`Abovesection`) OR (`ItemGroupKeys`.`Abovesection`=0))
          AND ((`ItemGroupKeys`.`Section`!=0 AND `ItemGroupKeys`.`Section`=NEW.`Section`) OR (`ItemGroupKeys`.`Section`=0))
          AND ((`ItemGroupKeys`.`SubSection`!=0 AND `ItemGroupKeys`.`SubSection`=NEW.`SubSection`) OR (`ItemGroupKeys`.`SubSection`=0))
          AND ((`ItemGroupKeys`.`Family`!=0 AND `ItemGroupKeys`.`Family`=NEW.`Family`) OR (`ItemGroupKeys`.`Family`=0))
          AND ((`ItemGroupKeys`.`SubFamily`!=0 AND `ItemGroupKeys`.`SubFamily`=NEW.`SubFamily`) OR (`ItemGroupKeys`.`SubFamily`=0))
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;

    WHEN 5 THEN
      IF ((NEW.`Code`!=0) AND (NEW.`OnKey`=TRUE)) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),`LocalTreeBoard`.`Board`, (`CFKeys`.`KeyCode`*10000)+(IF(`CFKeys`.`Level3`=0, `CFKeys`.`Level2`, `CFKeys`.`Level3`)), `CFKeys`.`Item` FROM `CFKeys`
          LEFT JOIN `LocalTreeBoard` ON `LocalTreeBoard`.`Level1`=`CFKeys`.`Level1` AND `LocalTreeBoard`.`Level2`=IF(`CFKeys`.`Level3`=0,0,`CFKeys`.`Level2`)
          WHERE `CFKeys`.`Item`=NEW.`Code` AND `CFKeys`.`Keyboard`=GLOBAL_GET("PLUload-kbd_db") AND (`CFKeys`.`EAN`!='00000000000000' AND `CFKeys`.`Item`!=0) AND Board IS NOT NULL
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;




    ELSE BEGIN END;
  END CASE;

  SET @toffer := IFNULL (GLOBAL_GET("table-offer"),0);
  SET @thabitual := IFNULL (GLOBAL_GET("table-habitual"),0);
  IF ((NEW.`Code`!=0) AND (NEW.`PLU`!=0)) THEN
    IF ((NEW.`Offer`=1) AND (@toffer!=0)) THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (GLOBAL_GET("current_group"),@toffer,NEW.`PLU`,NEW.`Code`)
        ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
    ELSE
      IF (@thabitual!=0) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),@thabitual,`ItemHabitual`.`Position`,NEW.`Code` FROM `ItemHabitual`
          WHERE `ItemHabitual`.`Code`=NEW.`Code`
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemsUpdated` AFTER UPDATE ON `Items`
FOR EACH ROW BEGIN
  IF(NEW.`EAN13`!=OLD.`EAN13`) THEN
    IF(NEW.`EAN13`!='') THEN
      INSERT INTO ItemsEAN13 (`Code`, `EAN13`) VALUES (NEW.`Code`, NEW.`EAN13`) ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
    END IF;
    IF(OLD.`EAN13`!='') THEN
      DELETE FROM ItemsEAN13 WHERE `Code`=OLD.`Code` AND `EAN13`=OLD.`EAN13`;
    END IF;
  END IF;
  IF GLOBAL_GET("PLUload-func") IS NOT NULL THEN
    CASE GLOBAL_GET("PLUload-func")

      WHEN 0 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,`Code` FROM `PLUKeys`
          WHERE `Code`=NEW.`Code` AND (`PLUKeys`.`Section`=GLOBAL_GET("current_group") OR (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard")) )
          ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;

      WHEN 1 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,`Code` FROM `PLUKeys`
          WHERE `Code`=NEW.`Code` AND (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard"))
          ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;
        IF(NEW.`Family`!=OLD.`Family`) THEN
          INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,OLD.`Family`,`PLUKey`,0 FROM `PLUKeys`
            WHERE `Code`=NEW.`Code` AND `PLUKeys`.`Section`=GLOBAL_GET("current_group")
            ON DUPLICATE KEY UPDATE `Code`=0;
        END IF;

        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,NEW.`Family`,`PLUKey`,`Code` FROM `PLUKeys`
          WHERE `Code`=NEW.`Code` AND `PLUKeys`.`Section`=GLOBAL_GET("current_group")
          ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;

      WHEN 2 THEN
        IF(OLD.`Family`!=NEW.`Family`) THEN
          SET @famor := (SELECT `FamilyIndex` FROM `CounterFamilies` WHERE `CounterCode`=GLOBAL_GET("current_counter") AND `FamilyCode`=OLD.`Family`);
          IF(@famor IS NOT NULL) THEN
            INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES(GLOBAL_GET("current_group"),@famor,OLD.`FamilyOrder`,0)
            ON DUPLICATE KEY UPDATE `Code`=0;
          END IF;
        END IF;
        SET @famor := (SELECT `FamilyIndex` FROM `CounterFamilies` WHERE `CounterCode`=GLOBAL_GET("current_counter") AND `FamilyCode`=NEW.`Family`);
        IF (@famor IS NOT NULL) THEN
          INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES(GLOBAL_GET("current_group"),@famor,NEW.`FamilyOrder`,
           IF(NEW.`OnKey`=TRUE, NEW.`Code`, 0))
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
        END IF;

        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,IF(NEW.`OnKey`=TRUE, `Code`, 0) FROM `PLUKeys`
          WHERE `Code`=NEW.`Code` AND (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard"))
          ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;

      WHEN 3 THEN
        IF (OLD.`Counter`!=NEW.`Counter`) AND (OLD.`Counter`=GLOBAL_GET("current_counter")) THEN
          INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (OLD.`Counter`,0,OLD.`PLU`, 0)
          ON DUPLICATE KEY UPDATE `Code`=0;
        END IF;
        IF (NEW.`Counter`!=0) AND (NEW.`PLU`!=0) AND (NEW.`Counter`=GLOBAL_GET("current_counter")) THEN
          INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (NEW.`Counter`,0,NEW.`PLU`, NEW.`Code`)
            ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
        END IF;

      WHEN 100 THEN
         IF (OLD.`Counter`!=NEW.`Counter`) THEN
          IF (OLD.`Counter`!=0) AND (OLD.`PLU`!=0) THEN
            INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT OLD.`Counter`,board_idx,OLD.`PLU`, 0
              FROM (SELECT IF(Code=GLOBAL_GET("current_counter"),1,@a:=@a+1) AS board_idx, Code FROM (SELECT @a:=1) AS ct, Counters) AS tt100 WHERE OLD.`Counter`=tt100.Code
           ON DUPLICATE KEY UPDATE `Code`=0;
          END IF;
          IF (NEW.`Counter`!=0) AND (NEW.`PLU`!=0) THEN
            INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Counter`,board_idx,NEW.`PLU`, NEW.`Code`
               FROM (SELECT IF(Code=GLOBAL_GET("current_counter"),1,@a:=@a+1) AS board_idx, Code FROM (SELECT @a:=1) AS ct, Counters) AS tt100 WHERE NEW.`Counter`=tt100.Code
            ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
          END IF;
        ELSE
          IF (NEW.`Counter`!=0) THEN
             IF (OLD.`PLU`!=0) THEN
              INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Counter`,board_idx,OLD.`PLU`, 0
                FROM (SELECT IF(Code=GLOBAL_GET("current_counter"),1,@a:=@a+1) AS board_idx, Code FROM (SELECT @a:=1) AS ct, Counters) AS tt100 WHERE NEW.`Counter`=tt100.Code
              ON DUPLICATE KEY UPDATE `Code`=`LocalPLUKeysChanges`.`Code`;
            END IF;
            IF (NEW.`PLU`!=0) THEN
              INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Counter`,board_idx,NEW.`PLU`, NEW.`Code`
                FROM (SELECT IF(Code=GLOBAL_GET("current_counter"),1,@a:=@a+1) AS board_idx, Code FROM (SELECT @a:=1) AS ct, Counters) AS tt100 WHERE NEW.`Counter`=tt100.Code
              ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
            END IF;
          END IF;
        END IF;

    WHEN 101 THEN
         IF (OLD.`Counter`!=NEW.`Counter`) THEN
          IF (OLD.`Counter`!=0) AND (OLD.`PLU`!=0) THEN
            INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT OLD.`Counter`,board_idx,OLD.`PLU`, 0
              FROM (SELECT @a:=@a+1 AS board_idx, Code FROM (SELECT @a:=0) AS ct, Counters) AS tt100 WHERE OLD.`Counter`=tt100.Code
           ON DUPLICATE KEY UPDATE `Code`=0;
          END IF;
          IF (NEW.`Counter`!=0) AND (NEW.`PLU`!=0) THEN
            INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Counter`,board_idx,NEW.`PLU`, NEW.`Code`
               FROM (SELECT @a:=@a+1 AS board_idx, Code FROM (SELECT @a:=0) AS ct, Counters) AS tt100 WHERE NEW.`Counter`=tt100.Code
            ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
          END IF;
        ELSE
          IF (NEW.`Counter`!=0) THEN
             IF (OLD.`PLU`!=0) THEN
              INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Counter`,board_idx,OLD.`PLU`, 0
                FROM (SELECT @a:=@a+1 AS board_idx, Code FROM (SELECT @a:=0) AS ct, Counters) AS tt100 WHERE NEW.`Counter`=tt100.Code
              ON DUPLICATE KEY UPDATE `Code`=`LocalPLUKeysChanges`.`Code`;
            END IF;
            IF (NEW.`PLU`!=0) THEN
              INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Counter`,board_idx,NEW.`PLU`, NEW.`Code`
                FROM (SELECT @a:=@a+1 AS board_idx, Code FROM (SELECT @a:=0) AS ct, Counters) AS tt100 WHERE NEW.`Counter`=tt100.Code
              ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
            END IF;
          END IF;
        END IF;

        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,IF(NEW.`OnKey`=TRUE, `Code`, 0) FROM `PLUKeys`
          WHERE `Code`=NEW.`Code` AND (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard"))
          ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;

    WHEN 4 THEN
      IF ((OLD.`PLU`!=0) AND (OLD.`Code`!=0)) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),`ItemGroupKeys`.`NumKey`,OLD.`PLU`,0 FROM `ItemGroupKeys`
          WHERE ((`ItemGroupKeys`.`Abovesection`!=0 AND `ItemGroupKeys`.`Abovesection`=OLD.`Abovesection`) OR (`ItemGroupKeys`.`Abovesection`=0))
          AND ((`ItemGroupKeys`.`Section`!=0 AND `ItemGroupKeys`.`Section`=OLD.`Section`) OR (`ItemGroupKeys`.`Section`=0))
          AND ((`ItemGroupKeys`.`SubSection`!=0 AND `ItemGroupKeys`.`SubSection`=OLD.`SubSection`) OR (`ItemGroupKeys`.`SubSection`=0))
          AND ((`ItemGroupKeys`.`Family`!=0 AND `ItemGroupKeys`.`Family`=OLD.`Family`) OR (`ItemGroupKeys`.`Family`=0))
          AND ((`ItemGroupKeys`.`SubFamily`!=0 AND `ItemGroupKeys`.`SubFamily`=OLD.`SubFamily`) OR (`ItemGroupKeys`.`SubFamily`=0))
          ON DUPLICATE KEY UPDATE `Code`=0;
      END IF;
      IF ((NEW.`PLU`!=0) AND (NEW.`Code`!=0)) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),`ItemGroupKeys`.`NumKey`,NEW.`PLU`,NEW.`Code` FROM `ItemGroupKeys`
          WHERE ((`ItemGroupKeys`.`Abovesection`!=0 AND `ItemGroupKeys`.`Abovesection`=NEW.`Abovesection`) OR (`ItemGroupKeys`.`Abovesection`=0))
          AND ((`ItemGroupKeys`.`Section`!=0 AND `ItemGroupKeys`.`Section`=NEW.`Section`) OR (`ItemGroupKeys`.`Section`=0))
          AND ((`ItemGroupKeys`.`SubSection`!=0 AND `ItemGroupKeys`.`SubSection`=NEW.`SubSection`) OR (`ItemGroupKeys`.`SubSection`=0))
          AND ((`ItemGroupKeys`.`Family`!=0 AND `ItemGroupKeys`.`Family`=NEW.`Family`) OR (`ItemGroupKeys`.`Family`=0))
          AND ((`ItemGroupKeys`.`SubFamily`!=0 AND `ItemGroupKeys`.`SubFamily`=NEW.`SubFamily`) OR (`ItemGroupKeys`.`SubFamily`=0))
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
     END IF;

    WHEN 5 THEN
      IF ((OLD.`Code`!=NEW.`Code`)  AND (OLD.`OnKey`=TRUE)) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),`LocalTreeBoard`.`Board`, (`CFKeys`.`KeyCode`*10000)+(IF(`CFKeys`.`Level3`=0, `CFKeys`.`Level2`, `CFKeys`.`Level3`)), 0 FROM `CFKeys`
          LEFT JOIN `LocalTreeBoard` ON `LocalTreeBoard`.`Level1`=`CFKeys`.`Level1` AND `LocalTreeBoard`.`Level2`=IF(`CFKeys`.`Level3`=0,0,`CFKeys`.`Level2`)
          WHERE `CFKeys`.`Item`=OLD.`Code` AND `CFKeys`.`Keyboard`=GLOBAL_GET("PLUload-kbd_db") AND (`CFKeys`.`EAN`!='00000000000000' AND `CFKeys`.`Item`!=0) AND Board IS NOT NULL
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;
      IF (NEW.`Code`!=0) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),`LocalTreeBoard`.`Board`, (`CFKeys`.`KeyCode`*10000)+(IF(`CFKeys`.`Level3`=0, `CFKeys`.`Level2`, `CFKeys`.`Level3`)), IF(NEW.`OnKey`=TRUE, NEW.`Code`, 0) FROM `CFKeys`
          LEFT JOIN `LocalTreeBoard` ON `LocalTreeBoard`.`Level1`=`CFKeys`.`Level1` AND `LocalTreeBoard`.`Level2`=IF(`CFKeys`.`Level3`=0,0,`CFKeys`.`Level2`)
          WHERE `CFKeys`.`Item`=NEW.`Code` AND `CFKeys`.`Keyboard`=GLOBAL_GET("PLUload-kbd_db") AND (`CFKeys`.`EAN`!='00000000000000' AND `CFKeys`.`Item`!=0) AND Board IS NOT NULL
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;




      ELSE BEGIN END;
    END CASE;
  END IF;

  SET @toffer := IFNULL (GLOBAL_GET("table-offer"),0);
  SET @thabitual := IFNULL (GLOBAL_GET("table-habitual"),0);
  IF ((OLD.`Code`!=0) AND (OLD.`PLU`!=0)) THEN
    IF ((OLD.`Offer`=1) AND (@toffer!=0)) THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (GLOBAL_GET("current_group"),@toffer,OLD.`PLU`,0)
        ON DUPLICATE KEY UPDATE `Code`=0;
    ELSE
      IF (@thabitual!=0) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),@thabitual,`ItemHabitual`.`Position`,0 FROM `ItemHabitual`
          WHERE `ItemHabitual`.`Code`=OLD.`Code`
          ON DUPLICATE KEY UPDATE `Code`=0;
      END IF;
    END IF;
  END IF;
  IF ((NEW.`Code`!=0) AND (NEW.`PLU`!=0)) THEN
    IF ((NEW.`Offer`=1) AND (@toffer!=0)) THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (GLOBAL_GET("current_group"),@toffer,NEW.`PLU`,NEW.`Code`)
        ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
    ELSE
      IF (@thabitual!=0) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),@thabitual,`ItemHabitual`.`Position`,NEW.`Code` FROM `ItemHabitual`
          WHERE `ItemHabitual`.`Code`=NEW.`Code`
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemsDeleted` BEFORE DELETE ON `Items`
FOR EACH ROW BEGIN
  SET @loadfunc := IFNULL (GLOBAL_GET("PLUload-func"),99);
  CASE @loadfunc

    WHEN 0 THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,0 FROM `PLUKeys`
        WHERE `Code`=OLD.`Code` AND (`PLUKeys`.`Section`=GLOBAL_GET("current_group") OR (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard")) )
        ON DUPLICATE KEY UPDATE `Code`=0;

    WHEN 1 THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,0 FROM `PLUKeys`
        WHERE `Code`=OLD.`Code` AND (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard"))
        ON DUPLICATE KEY UPDATE `Code`=0;
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,OLD.`Family`,`PLUKey`,0 FROM `PLUKeys`
        WHERE `Code`=OLD.`Code` AND `PLUKeys`.`Section`=GLOBAL_GET("current_group")
        ON DUPLICATE KEY UPDATE `Code`=0;

    WHEN 2 THEN
      SET @famor := (SELECT `FamilyIndex` FROM `CounterFamilies` WHERE `CounterCode`=GLOBAL_GET("current_counter") AND `FamilyCode`=OLD.`Family`);
      IF (@famor IS NOT NULL) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES(GLOBAL_GET("current_group"),@famor,OLD.`FamilyOrder`,0)
        ON DUPLICATE KEY UPDATE `Code`=0;
      END IF;
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT `Section`,`Board`,`PLUKey`,0 FROM `PLUKeys`
        WHERE `Code`=OLD.`Code` AND (`PLUKeys`.`Section`=0 AND `PLUKeys`.`Board`=GLOBAL_GET("current_mupboard"))
        ON DUPLICATE KEY UPDATE `Code`=0;

     WHEN 3 THEN
      IF (OLD.`Counter`!=0) AND (OLD.`PLU`!=0) AND (OLD.`Counter`=GLOBAL_GET("current_counter")) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (OLD.`Counter`,0,OLD.`PLU`, 0)
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;

      WHEN 100 THEN
        IF (OLD.`Counter`!=0) AND (OLD.`PLU`!=0) THEN
          INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT OLD.`Counter`,board_idx,OLD.`PLU`, 0
            FROM (SELECT IF(Code=GLOBAL_GET("current_counter"),1,@a:=@a+1) AS board_idx, Code FROM (SELECT @a:=1) AS ct, Counters) AS tt100 WHERE OLD.`Counter`=tt100.Code
          ON DUPLICATE KEY UPDATE `Code`=0;
        END IF;

    WHEN 101 THEN
        IF (OLD.`Counter`!=0) AND (OLD.`PLU`!=0) THEN
          INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT OLD.`Counter`,board_idx,OLD.`PLU`, 0
            FROM (SELECT @a:=@a+1 AS board_idx, Code FROM (SELECT @a:=0) AS ct, Counters) AS tt100 WHERE OLD.`Counter`=tt100.Code
          ON DUPLICATE KEY UPDATE `Code`=0;
        END IF;

    WHEN 4 THEN
      IF ((OLD.`PLU`!=0) AND (OLD.`Code`!=0)) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),`ItemGroupKeys`.`NumKey`,OLD.`PLU`,0 FROM `ItemGroupKeys`
          WHERE ((`ItemGroupKeys`.`Abovesection`!=0 AND `ItemGroupKeys`.`Abovesection`=OLD.`Abovesection`) OR (`ItemGroupKeys`.`Abovesection`=0))
          AND ((`ItemGroupKeys`.`Section`!=0 AND `ItemGroupKeys`.`Section`=OLD.`Section`) OR (`ItemGroupKeys`.`Section`=0))
          AND ((`ItemGroupKeys`.`SubSection`!=0 AND `ItemGroupKeys`.`SubSection`=OLD.`SubSection`) OR (`ItemGroupKeys`.`SubSection`=0))
          AND ((`ItemGroupKeys`.`Family`!=0 AND `ItemGroupKeys`.`Family`=OLD.`Family`) OR (`ItemGroupKeys`.`Family`=0))
          AND ((`ItemGroupKeys`.`SubFamily`!=0 AND `ItemGroupKeys`.`SubFamily`=OLD.`SubFamily`) OR (`ItemGroupKeys`.`SubFamily`=0))
          ON DUPLICATE KEY UPDATE `Code`=0;
      END IF;

    WHEN 5 THEN
      IF ((OLD.`Code`!=0) AND (OLD.`OnKey`=TRUE)) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),`LocalTreeBoard`.`Board`, (`CFKeys`.`KeyCode`*10000)+(IF(`CFKeys`.`Level3`=0, `CFKeys`.`Level2`, `CFKeys`.`Level3`)), 0 FROM `CFKeys`
          LEFT JOIN `LocalTreeBoard` ON `LocalTreeBoard`.`Level1`=`CFKeys`.`Level1` AND `LocalTreeBoard`.`Level2`=IF(`CFKeys`.`Level3`=0,0,`CFKeys`.`Level2`)
          WHERE `CFKeys`.`Item`=OLD.`Code` AND `CFKeys`.`Keyboard`=GLOBAL_GET("PLUload-kbd_db") AND (`CFKeys`.`EAN`!='00000000000000' AND `CFKeys`.`Item`!=0) AND Board IS NOT NULL
          ON DUPLICATE KEY UPDATE `Code`=VALUES(`Code`);
      END IF;




    ELSE BEGIN END;
  END CASE;

  SET @toffer := IFNULL (GLOBAL_GET("table-offer"),0);
  SET @thabitual := IFNULL (GLOBAL_GET("table-habitual"),0);
  IF ((OLD.`Code`!=0) AND (OLD.`PLU`!=0)) THEN
    IF ((OLD.`Offer`=1) AND (@toffer!=0)) THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (GLOBAL_GET("current_group"),@toffer,OLD.`PLU`,0)
        ON DUPLICATE KEY UPDATE `Code`=0;
    ELSE
      IF (@thabitual!=0) THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`)
          SELECT GLOBAL_GET("current_group"),@thabitual,`ItemHabitual`.`Position`,0 FROM `ItemHabitual` WHERE `ItemHabitual`.`Code`=OLD.`Code`
          ON DUPLICATE KEY UPDATE `Code`=0;
      END IF;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `itemschangepoints`
--

DROP TABLE IF EXISTS `itemschangepoints`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemschangepoints` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` int(11) NOT NULL,
  `Points` int(11) NOT NULL DEFAULT '0',
  `Quantity` decimal(5,3) NOT NULL DEFAULT '1.000',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Code` (`Code`),
  CONSTRAINT `icpitem` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemschangepoints`
--

LOCK TABLES `itemschangepoints` WRITE;
/*!40000 ALTER TABLE `itemschangepoints` DISABLE KEYS */;
/*!40000 ALTER TABLE `itemschangepoints` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `itemsean13`
--

DROP TABLE IF EXISTS `itemsean13`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemsean13` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` int(11) NOT NULL,
  `EAN13` decimal(13,0) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `EAN13` (`EAN13`),
  KEY `iean13_item` (`Code`),
  CONSTRAINT `iean13_item` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemsean13`
--

LOCK TABLES `itemsean13` WRITE;
/*!40000 ALTER TABLE `itemsean13` DISABLE KEYS */;
/*!40000 ALTER TABLE `itemsean13` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `itemsextrapoints`
--

DROP TABLE IF EXISTS `itemsextrapoints`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemsextrapoints` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` int(11) NOT NULL,
  `Points` int(11) NOT NULL DEFAULT '0',
  `Quantity` decimal(5,3) NOT NULL DEFAULT '1.000',
  `FromQuantity` decimal(5,3) NOT NULL DEFAULT '0.000',
  `Proportional` tinyint(1) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Code` (`Code`),
  CONSTRAINT `iepitem` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemsextrapoints`
--

LOCK TABLES `itemsextrapoints` WRITE;
/*!40000 ALTER TABLE `itemsextrapoints` DISABLE KEYS */;
/*!40000 ALTER TABLE `itemsextrapoints` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemsExtraPointsInserted` AFTER INSERT ON `ItemsExtraPoints`
FOR EACH ROW BEGIN
  UPDATE `Items` SET `Items`.`Code`=`Items`.`Code` WHERE `Items`.`Code`=NEW.`Code`;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemsExtraPointsUpdated` AFTER UPDATE ON `ItemsExtraPoints`
FOR EACH ROW BEGIN
  IF(OLD.`Code`!=NEW.`Code`) THEN
    UPDATE `Items` SET `Items`.`Code`=`Items`.`Code` WHERE `Items`.`Code`=OLD.`Code`;
  END IF;
  UPDATE `Items` SET `Items`.`Code`=`Items`.`Code` WHERE `Items`.`Code`=NEW.`Code`;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ItemsExtraPointsDeleted` AFTER DELETE ON `ItemsExtraPoints`
FOR EACH ROW BEGIN
  UPDATE `Items` SET `Items`.`Code`=`Items`.`Code` WHERE `Items`.`Code`=OLD.`Code`;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `itemtextbonus`
--

DROP TABLE IF EXISTS `itemtextbonus`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemtextbonus` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` int(11) NOT NULL,
  `NumLine` smallint(6) DEFAULT '1',
  `Text` varchar(40) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `itemtext_idx` (`Code`,`NumLine`),
  CONSTRAINT `itemtextbonus_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=272821 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemtextbonus`
--

LOCK TABLES `itemtextbonus` WRITE;
/*!40000 ALTER TABLE `itemtextbonus` DISABLE KEYS */;
INSERT INTO `itemtextbonus` VALUES (257241,1127,1,''),(257242,1127,2,''),(257243,1127,3,''),(257244,1127,4,''),(257245,1127,5,''),(257246,1127,6,''),(257247,1127,7,''),(257248,1127,8,''),(257249,1127,9,''),(257250,1127,10,''),(257251,1127,11,''),(257252,1127,12,''),(257253,1127,13,''),(257254,1127,14,''),(257255,1127,15,''),(257256,1127,16,''),(257257,1127,17,''),(257258,1127,18,''),(257259,1127,19,''),(257260,1127,20,''),(257261,1176,1,''),(257262,1176,2,''),(257263,1176,3,''),(257264,1176,4,''),(257265,1176,5,''),(257266,1176,6,''),(257267,1176,7,''),(257268,1176,8,''),(257269,1176,9,''),(257270,1176,10,''),(257271,1176,11,''),(257272,1176,12,''),(257273,1176,13,''),(257274,1176,14,''),(257275,1176,15,''),(257276,1176,16,''),(257277,1176,17,''),(257278,1176,18,''),(257279,1176,19,''),(257280,1176,20,''),(257281,1207,1,''),(257282,1207,2,''),(257283,1207,3,''),(257284,1207,4,''),(257285,1207,5,''),(257286,1207,6,''),(257287,1207,7,''),(257288,1207,8,''),(257289,1207,9,''),(257290,1207,10,''),(257291,1207,11,''),(257292,1207,12,''),(257293,1207,13,''),(257294,1207,14,''),(257295,1207,15,''),(257296,1207,16,''),(257297,1207,17,''),(257298,1207,18,''),(257299,1207,19,''),(257300,1207,20,''),(257301,1616,1,''),(257302,1616,2,''),(257303,1616,3,''),(257304,1616,4,''),(257305,1616,5,''),(257306,1616,6,''),(257307,1616,7,''),(257308,1616,8,''),(257309,1616,9,''),(257310,1616,10,''),(257311,1616,11,''),(257312,1616,12,''),(257313,1616,13,''),(257314,1616,14,''),(257315,1616,15,''),(257316,1616,16,''),(257317,1616,17,''),(257318,1616,18,''),(257319,1616,19,''),(257320,1616,20,''),(257321,1617,1,''),(257322,1617,2,''),(257323,1617,3,''),(257324,1617,4,''),(257325,1617,5,''),(257326,1617,6,''),(257327,1617,7,''),(257328,1617,8,''),(257329,1617,9,''),(257330,1617,10,''),(257331,1617,11,''),(257332,1617,12,''),(257333,1617,13,''),(257334,1617,14,''),(257335,1617,15,''),(257336,1617,16,''),(257337,1617,17,''),(257338,1617,18,''),(257339,1617,19,''),(257340,1617,20,''),(257341,1621,1,''),(257342,1621,2,''),(257343,1621,3,''),(257344,1621,4,''),(257345,1621,5,''),(257346,1621,6,''),(257347,1621,7,''),(257348,1621,8,''),(257349,1621,9,''),(257350,1621,10,''),(257351,1621,11,''),(257352,1621,12,''),(257353,1621,13,''),(257354,1621,14,''),(257355,1621,15,''),(257356,1621,16,''),(257357,1621,17,''),(257358,1621,18,''),(257359,1621,19,''),(257360,1621,20,''),(257361,1677,1,''),(257362,1677,2,''),(257363,1677,3,''),(257364,1677,4,''),(257365,1677,5,''),(257366,1677,6,''),(257367,1677,7,''),(257368,1677,8,''),(257369,1677,9,''),(257370,1677,10,''),(257371,1677,11,''),(257372,1677,12,''),(257373,1677,13,''),(257374,1677,14,''),(257375,1677,15,''),(257376,1677,16,''),(257377,1677,17,''),(257378,1677,18,''),(257379,1677,19,''),(257380,1677,20,''),(257381,2046,1,''),(257382,2046,2,''),(257383,2046,3,''),(257384,2046,4,''),(257385,2046,5,''),(257386,2046,6,''),(257387,2046,7,''),(257388,2046,8,''),(257389,2046,9,''),(257390,2046,10,''),(257391,2046,11,''),(257392,2046,12,''),(257393,2046,13,''),(257394,2046,14,''),(257395,2046,15,''),(257396,2046,16,''),(257397,2046,17,''),(257398,2046,18,''),(257399,2046,19,''),(257400,2046,20,''),(257401,2081,1,''),(257402,2081,2,''),(257403,2081,3,''),(257404,2081,4,''),(257405,2081,5,''),(257406,2081,6,''),(257407,2081,7,''),(257408,2081,8,''),(257409,2081,9,''),(257410,2081,10,''),(257411,2081,11,''),(257412,2081,12,''),(257413,2081,13,''),(257414,2081,14,''),(257415,2081,15,''),(257416,2081,16,''),(257417,2081,17,''),(257418,2081,18,''),(257419,2081,19,''),(257420,2081,20,''),(257421,2276,1,''),(257422,2276,2,''),(257423,2276,3,''),(257424,2276,4,''),(257425,2276,5,''),(257426,2276,6,''),(257427,2276,7,''),(257428,2276,8,''),(257429,2276,9,''),(257430,2276,10,''),(257431,2276,11,''),(257432,2276,12,''),(257433,2276,13,''),(257434,2276,14,''),(257435,2276,15,''),(257436,2276,16,''),(257437,2276,17,''),(257438,2276,18,''),(257439,2276,19,''),(257440,2276,20,''),(257441,2279,1,''),(257442,2279,2,''),(257443,2279,3,''),(257444,2279,4,''),(257445,2279,5,''),(257446,2279,6,''),(257447,2279,7,''),(257448,2279,8,''),(257449,2279,9,''),(257450,2279,10,''),(257451,2279,11,''),(257452,2279,12,''),(257453,2279,13,''),(257454,2279,14,''),(257455,2279,15,''),(257456,2279,16,''),(257457,2279,17,''),(257458,2279,18,''),(257459,2279,19,''),(257460,2279,20,''),(257461,2355,1,''),(257462,2355,2,''),(257463,2355,3,''),(257464,2355,4,''),(257465,2355,5,''),(257466,2355,6,''),(257467,2355,7,''),(257468,2355,8,''),(257469,2355,9,''),(257470,2355,10,''),(257471,2355,11,''),(257472,2355,12,''),(257473,2355,13,''),(257474,2355,14,''),(257475,2355,15,''),(257476,2355,16,''),(257477,2355,17,''),(257478,2355,18,''),(257479,2355,19,''),(257480,2355,20,''),(257481,2467,1,''),(257482,2467,2,''),(257483,2467,3,''),(257484,2467,4,''),(257485,2467,5,''),(257486,2467,6,''),(257487,2467,7,''),(257488,2467,8,''),(257489,2467,9,''),(257490,2467,10,''),(257491,2467,11,''),(257492,2467,12,''),(257493,2467,13,''),(257494,2467,14,''),(257495,2467,15,''),(257496,2467,16,''),(257497,2467,17,''),(257498,2467,18,''),(257499,2467,19,''),(257500,2467,20,''),(257501,2487,1,''),(257502,2487,2,''),(257503,2487,3,''),(257504,2487,4,''),(257505,2487,5,''),(257506,2487,6,''),(257507,2487,7,''),(257508,2487,8,''),(257509,2487,9,''),(257510,2487,10,''),(257511,2487,11,''),(257512,2487,12,''),(257513,2487,13,''),(257514,2487,14,''),(257515,2487,15,''),(257516,2487,16,''),(257517,2487,17,''),(257518,2487,18,''),(257519,2487,19,''),(257520,2487,20,''),(257521,2514,1,''),(257522,2514,2,''),(257523,2514,3,''),(257524,2514,4,''),(257525,2514,5,''),(257526,2514,6,''),(257527,2514,7,''),(257528,2514,8,''),(257529,2514,9,''),(257530,2514,10,''),(257531,2514,11,''),(257532,2514,12,''),(257533,2514,13,''),(257534,2514,14,''),(257535,2514,15,''),(257536,2514,16,''),(257537,2514,17,''),(257538,2514,18,''),(257539,2514,19,''),(257540,2514,20,''),(257541,2515,1,''),(257542,2515,2,''),(257543,2515,3,''),(257544,2515,4,''),(257545,2515,5,''),(257546,2515,6,''),(257547,2515,7,''),(257548,2515,8,''),(257549,2515,9,''),(257550,2515,10,''),(257551,2515,11,''),(257552,2515,12,''),(257553,2515,13,''),(257554,2515,14,''),(257555,2515,15,''),(257556,2515,16,''),(257557,2515,17,''),(257558,2515,18,''),(257559,2515,19,''),(257560,2515,20,''),(257561,2525,1,''),(257562,2525,2,''),(257563,2525,3,''),(257564,2525,4,''),(257565,2525,5,''),(257566,2525,6,''),(257567,2525,7,''),(257568,2525,8,''),(257569,2525,9,''),(257570,2525,10,''),(257571,2525,11,''),(257572,2525,12,''),(257573,2525,13,''),(257574,2525,14,''),(257575,2525,15,''),(257576,2525,16,''),(257577,2525,17,''),(257578,2525,18,''),(257579,2525,19,''),(257580,2525,20,''),(257581,2537,1,''),(257582,2537,2,''),(257583,2537,3,''),(257584,2537,4,''),(257585,2537,5,''),(257586,2537,6,''),(257587,2537,7,''),(257588,2537,8,''),(257589,2537,9,''),(257590,2537,10,''),(257591,2537,11,''),(257592,2537,12,''),(257593,2537,13,''),(257594,2537,14,''),(257595,2537,15,''),(257596,2537,16,''),(257597,2537,17,''),(257598,2537,18,''),(257599,2537,19,''),(257600,2537,20,''),(257601,2550,1,''),(257602,2550,2,''),(257603,2550,3,''),(257604,2550,4,''),(257605,2550,5,''),(257606,2550,6,''),(257607,2550,7,''),(257608,2550,8,''),(257609,2550,9,''),(257610,2550,10,''),(257611,2550,11,''),(257612,2550,12,''),(257613,2550,13,''),(257614,2550,14,''),(257615,2550,15,''),(257616,2550,16,''),(257617,2550,17,''),(257618,2550,18,''),(257619,2550,19,''),(257620,2550,20,''),(257621,2607,1,''),(257622,2607,2,''),(257623,2607,3,''),(257624,2607,4,''),(257625,2607,5,''),(257626,2607,6,''),(257627,2607,7,''),(257628,2607,8,''),(257629,2607,9,''),(257630,2607,10,''),(257631,2607,11,''),(257632,2607,12,''),(257633,2607,13,''),(257634,2607,14,''),(257635,2607,15,''),(257636,2607,16,''),(257637,2607,17,''),(257638,2607,18,''),(257639,2607,19,''),(257640,2607,20,''),(257641,2655,1,''),(257642,2655,2,''),(257643,2655,3,''),(257644,2655,4,''),(257645,2655,5,''),(257646,2655,6,''),(257647,2655,7,''),(257648,2655,8,''),(257649,2655,9,''),(257650,2655,10,''),(257651,2655,11,''),(257652,2655,12,''),(257653,2655,13,''),(257654,2655,14,''),(257655,2655,15,''),(257656,2655,16,''),(257657,2655,17,''),(257658,2655,18,''),(257659,2655,19,''),(257660,2655,20,''),(257661,2696,1,''),(257662,2696,2,''),(257663,2696,3,''),(257664,2696,4,''),(257665,2696,5,''),(257666,2696,6,''),(257667,2696,7,''),(257668,2696,8,''),(257669,2696,9,''),(257670,2696,10,''),(257671,2696,11,''),(257672,2696,12,''),(257673,2696,13,''),(257674,2696,14,''),(257675,2696,15,''),(257676,2696,16,''),(257677,2696,17,''),(257678,2696,18,''),(257679,2696,19,''),(257680,2696,20,''),(257681,2714,1,''),(257682,2714,2,''),(257683,2714,3,''),(257684,2714,4,''),(257685,2714,5,''),(257686,2714,6,''),(257687,2714,7,''),(257688,2714,8,''),(257689,2714,9,''),(257690,2714,10,''),(257691,2714,11,''),(257692,2714,12,''),(257693,2714,13,''),(257694,2714,14,''),(257695,2714,15,''),(257696,2714,16,''),(257697,2714,17,''),(257698,2714,18,''),(257699,2714,19,''),(257700,2714,20,''),(257701,2716,1,''),(257702,2716,2,''),(257703,2716,3,''),(257704,2716,4,''),(257705,2716,5,''),(257706,2716,6,''),(257707,2716,7,''),(257708,2716,8,''),(257709,2716,9,''),(257710,2716,10,''),(257711,2716,11,''),(257712,2716,12,''),(257713,2716,13,''),(257714,2716,14,''),(257715,2716,15,''),(257716,2716,16,''),(257717,2716,17,''),(257718,2716,18,''),(257719,2716,19,''),(257720,2716,20,''),(257721,2722,1,''),(257722,2722,2,''),(257723,2722,3,''),(257724,2722,4,''),(257725,2722,5,''),(257726,2722,6,''),(257727,2722,7,''),(257728,2722,8,''),(257729,2722,9,''),(257730,2722,10,''),(257731,2722,11,''),(257732,2722,12,''),(257733,2722,13,''),(257734,2722,14,''),(257735,2722,15,''),(257736,2722,16,''),(257737,2722,17,''),(257738,2722,18,''),(257739,2722,19,''),(257740,2722,20,''),(257741,2745,1,''),(257742,2745,2,''),(257743,2745,3,''),(257744,2745,4,''),(257745,2745,5,''),(257746,2745,6,''),(257747,2745,7,''),(257748,2745,8,''),(257749,2745,9,''),(257750,2745,10,''),(257751,2745,11,''),(257752,2745,12,''),(257753,2745,13,''),(257754,2745,14,''),(257755,2745,15,''),(257756,2745,16,''),(257757,2745,17,''),(257758,2745,18,''),(257759,2745,19,''),(257760,2745,20,''),(257761,2938,1,''),(257762,2938,2,''),(257763,2938,3,''),(257764,2938,4,''),(257765,2938,5,''),(257766,2938,6,''),(257767,2938,7,''),(257768,2938,8,''),(257769,2938,9,''),(257770,2938,10,''),(257771,2938,11,''),(257772,2938,12,''),(257773,2938,13,''),(257774,2938,14,''),(257775,2938,15,''),(257776,2938,16,''),(257777,2938,17,''),(257778,2938,18,''),(257779,2938,19,''),(257780,2938,20,''),(257781,2987,1,''),(257782,2987,2,''),(257783,2987,3,''),(257784,2987,4,''),(257785,2987,5,''),(257786,2987,6,''),(257787,2987,7,''),(257788,2987,8,''),(257789,2987,9,''),(257790,2987,10,''),(257791,2987,11,''),(257792,2987,12,''),(257793,2987,13,''),(257794,2987,14,''),(257795,2987,15,''),(257796,2987,16,''),(257797,2987,17,''),(257798,2987,18,''),(257799,2987,19,''),(257800,2987,20,''),(257801,3155,1,''),(257802,3155,2,''),(257803,3155,3,''),(257804,3155,4,''),(257805,3155,5,''),(257806,3155,6,''),(257807,3155,7,''),(257808,3155,8,''),(257809,3155,9,''),(257810,3155,10,''),(257811,3155,11,''),(257812,3155,12,''),(257813,3155,13,''),(257814,3155,14,''),(257815,3155,15,''),(257816,3155,16,''),(257817,3155,17,''),(257818,3155,18,''),(257819,3155,19,''),(257820,3155,20,''),(257821,3195,1,''),(257822,3195,2,''),(257823,3195,3,''),(257824,3195,4,''),(257825,3195,5,''),(257826,3195,6,''),(257827,3195,7,''),(257828,3195,8,''),(257829,3195,9,''),(257830,3195,10,''),(257831,3195,11,''),(257832,3195,12,''),(257833,3195,13,''),(257834,3195,14,''),(257835,3195,15,''),(257836,3195,16,''),(257837,3195,17,''),(257838,3195,18,''),(257839,3195,19,''),(257840,3195,20,''),(257841,3276,1,''),(257842,3276,2,''),(257843,3276,3,''),(257844,3276,4,''),(257845,3276,5,''),(257846,3276,6,''),(257847,3276,7,''),(257848,3276,8,''),(257849,3276,9,''),(257850,3276,10,''),(257851,3276,11,''),(257852,3276,12,''),(257853,3276,13,''),(257854,3276,14,''),(257855,3276,15,''),(257856,3276,16,''),(257857,3276,17,''),(257858,3276,18,''),(257859,3276,19,''),(257860,3276,20,''),(257861,3277,1,''),(257862,3277,2,''),(257863,3277,3,''),(257864,3277,4,''),(257865,3277,5,''),(257866,3277,6,''),(257867,3277,7,''),(257868,3277,8,''),(257869,3277,9,''),(257870,3277,10,''),(257871,3277,11,''),(257872,3277,12,''),(257873,3277,13,''),(257874,3277,14,''),(257875,3277,15,''),(257876,3277,16,''),(257877,3277,17,''),(257878,3277,18,''),(257879,3277,19,''),(257880,3277,20,''),(257881,3406,1,''),(257882,3406,2,''),(257883,3406,3,''),(257884,3406,4,''),(257885,3406,5,''),(257886,3406,6,''),(257887,3406,7,''),(257888,3406,8,''),(257889,3406,9,''),(257890,3406,10,''),(257891,3406,11,''),(257892,3406,12,''),(257893,3406,13,''),(257894,3406,14,''),(257895,3406,15,''),(257896,3406,16,''),(257897,3406,17,''),(257898,3406,18,''),(257899,3406,19,''),(257900,3406,20,''),(257901,3481,1,''),(257902,3481,2,''),(257903,3481,3,''),(257904,3481,4,''),(257905,3481,5,''),(257906,3481,6,''),(257907,3481,7,''),(257908,3481,8,''),(257909,3481,9,''),(257910,3481,10,''),(257911,3481,11,''),(257912,3481,12,''),(257913,3481,13,''),(257914,3481,14,''),(257915,3481,15,''),(257916,3481,16,''),(257917,3481,17,''),(257918,3481,18,''),(257919,3481,19,''),(257920,3481,20,''),(257921,3628,1,''),(257922,3628,2,''),(257923,3628,3,''),(257924,3628,4,''),(257925,3628,5,''),(257926,3628,6,''),(257927,3628,7,''),(257928,3628,8,''),(257929,3628,9,''),(257930,3628,10,''),(257931,3628,11,''),(257932,3628,12,''),(257933,3628,13,''),(257934,3628,14,''),(257935,3628,15,''),(257936,3628,16,''),(257937,3628,17,''),(257938,3628,18,''),(257939,3628,19,''),(257940,3628,20,''),(257941,3630,1,''),(257942,3630,2,''),(257943,3630,3,''),(257944,3630,4,''),(257945,3630,5,''),(257946,3630,6,''),(257947,3630,7,''),(257948,3630,8,''),(257949,3630,9,''),(257950,3630,10,''),(257951,3630,11,''),(257952,3630,12,''),(257953,3630,13,''),(257954,3630,14,''),(257955,3630,15,''),(257956,3630,16,''),(257957,3630,17,''),(257958,3630,18,''),(257959,3630,19,''),(257960,3630,20,''),(257961,3637,1,''),(257962,3637,2,''),(257963,3637,3,''),(257964,3637,4,''),(257965,3637,5,''),(257966,3637,6,''),(257967,3637,7,''),(257968,3637,8,''),(257969,3637,9,''),(257970,3637,10,''),(257971,3637,11,''),(257972,3637,12,''),(257973,3637,13,''),(257974,3637,14,''),(257975,3637,15,''),(257976,3637,16,''),(257977,3637,17,''),(257978,3637,18,''),(257979,3637,19,''),(257980,3637,20,''),(257981,3689,1,''),(257982,3689,2,''),(257983,3689,3,''),(257984,3689,4,''),(257985,3689,5,''),(257986,3689,6,''),(257987,3689,7,''),(257988,3689,8,''),(257989,3689,9,''),(257990,3689,10,''),(257991,3689,11,''),(257992,3689,12,''),(257993,3689,13,''),(257994,3689,14,''),(257995,3689,15,''),(257996,3689,16,''),(257997,3689,17,''),(257998,3689,18,''),(257999,3689,19,''),(258000,3689,20,''),(258001,3763,1,''),(258002,3763,2,''),(258003,3763,3,''),(258004,3763,4,''),(258005,3763,5,''),(258006,3763,6,''),(258007,3763,7,''),(258008,3763,8,''),(258009,3763,9,''),(258010,3763,10,''),(258011,3763,11,''),(258012,3763,12,''),(258013,3763,13,''),(258014,3763,14,''),(258015,3763,15,''),(258016,3763,16,''),(258017,3763,17,''),(258018,3763,18,''),(258019,3763,19,''),(258020,3763,20,''),(258021,3776,1,''),(258022,3776,2,''),(258023,3776,3,''),(258024,3776,4,''),(258025,3776,5,''),(258026,3776,6,''),(258027,3776,7,''),(258028,3776,8,''),(258029,3776,9,''),(258030,3776,10,''),(258031,3776,11,''),(258032,3776,12,''),(258033,3776,13,''),(258034,3776,14,''),(258035,3776,15,''),(258036,3776,16,''),(258037,3776,17,''),(258038,3776,18,''),(258039,3776,19,''),(258040,3776,20,''),(258041,3803,1,''),(258042,3803,2,''),(258043,3803,3,''),(258044,3803,4,''),(258045,3803,5,''),(258046,3803,6,''),(258047,3803,7,''),(258048,3803,8,''),(258049,3803,9,''),(258050,3803,10,''),(258051,3803,11,''),(258052,3803,12,''),(258053,3803,13,''),(258054,3803,14,''),(258055,3803,15,''),(258056,3803,16,''),(258057,3803,17,''),(258058,3803,18,''),(258059,3803,19,''),(258060,3803,20,''),(258061,3806,1,''),(258062,3806,2,''),(258063,3806,3,''),(258064,3806,4,''),(258065,3806,5,''),(258066,3806,6,''),(258067,3806,7,''),(258068,3806,8,''),(258069,3806,9,''),(258070,3806,10,''),(258071,3806,11,''),(258072,3806,12,''),(258073,3806,13,''),(258074,3806,14,''),(258075,3806,15,''),(258076,3806,16,''),(258077,3806,17,''),(258078,3806,18,''),(258079,3806,19,''),(258080,3806,20,''),(258081,3830,1,''),(258082,3830,2,''),(258083,3830,3,''),(258084,3830,4,''),(258085,3830,5,''),(258086,3830,6,''),(258087,3830,7,''),(258088,3830,8,''),(258089,3830,9,''),(258090,3830,10,''),(258091,3830,11,''),(258092,3830,12,''),(258093,3830,13,''),(258094,3830,14,''),(258095,3830,15,''),(258096,3830,16,''),(258097,3830,17,''),(258098,3830,18,''),(258099,3830,19,''),(258100,3830,20,''),(258101,3884,1,''),(258102,3884,2,''),(258103,3884,3,''),(258104,3884,4,''),(258105,3884,5,''),(258106,3884,6,''),(258107,3884,7,''),(258108,3884,8,''),(258109,3884,9,''),(258110,3884,10,''),(258111,3884,11,''),(258112,3884,12,''),(258113,3884,13,''),(258114,3884,14,''),(258115,3884,15,''),(258116,3884,16,''),(258117,3884,17,''),(258118,3884,18,''),(258119,3884,19,''),(258120,3884,20,''),(258121,3888,1,''),(258122,3888,2,''),(258123,3888,3,''),(258124,3888,4,''),(258125,3888,5,''),(258126,3888,6,''),(258127,3888,7,''),(258128,3888,8,''),(258129,3888,9,''),(258130,3888,10,''),(258131,3888,11,''),(258132,3888,12,''),(258133,3888,13,''),(258134,3888,14,''),(258135,3888,15,''),(258136,3888,16,''),(258137,3888,17,''),(258138,3888,18,''),(258139,3888,19,''),(258140,3888,20,''),(258141,3889,1,''),(258142,3889,2,''),(258143,3889,3,''),(258144,3889,4,''),(258145,3889,5,''),(258146,3889,6,''),(258147,3889,7,''),(258148,3889,8,''),(258149,3889,9,''),(258150,3889,10,''),(258151,3889,11,''),(258152,3889,12,''),(258153,3889,13,''),(258154,3889,14,''),(258155,3889,15,''),(258156,3889,16,''),(258157,3889,17,''),(258158,3889,18,''),(258159,3889,19,''),(258160,3889,20,''),(258161,3907,1,''),(258162,3907,2,''),(258163,3907,3,''),(258164,3907,4,''),(258165,3907,5,''),(258166,3907,6,''),(258167,3907,7,''),(258168,3907,8,''),(258169,3907,9,''),(258170,3907,10,''),(258171,3907,11,''),(258172,3907,12,''),(258173,3907,13,''),(258174,3907,14,''),(258175,3907,15,''),(258176,3907,16,''),(258177,3907,17,''),(258178,3907,18,''),(258179,3907,19,''),(258180,3907,20,''),(258181,3914,1,''),(258182,3914,2,''),(258183,3914,3,''),(258184,3914,4,''),(258185,3914,5,''),(258186,3914,6,''),(258187,3914,7,''),(258188,3914,8,''),(258189,3914,9,''),(258190,3914,10,''),(258191,3914,11,''),(258192,3914,12,''),(258193,3914,13,''),(258194,3914,14,''),(258195,3914,15,''),(258196,3914,16,''),(258197,3914,17,''),(258198,3914,18,''),(258199,3914,19,''),(258200,3914,20,''),(258201,3922,1,''),(258202,3922,2,''),(258203,3922,3,''),(258204,3922,4,''),(258205,3922,5,''),(258206,3922,6,''),(258207,3922,7,''),(258208,3922,8,''),(258209,3922,9,''),(258210,3922,10,''),(258211,3922,11,''),(258212,3922,12,''),(258213,3922,13,''),(258214,3922,14,''),(258215,3922,15,''),(258216,3922,16,''),(258217,3922,17,''),(258218,3922,18,''),(258219,3922,19,''),(258220,3922,20,''),(258221,3929,1,''),(258222,3929,2,''),(258223,3929,3,''),(258224,3929,4,''),(258225,3929,5,''),(258226,3929,6,''),(258227,3929,7,''),(258228,3929,8,''),(258229,3929,9,''),(258230,3929,10,''),(258231,3929,11,''),(258232,3929,12,''),(258233,3929,13,''),(258234,3929,14,''),(258235,3929,15,''),(258236,3929,16,''),(258237,3929,17,''),(258238,3929,18,''),(258239,3929,19,''),(258240,3929,20,''),(258241,3932,1,''),(258242,3932,2,''),(258243,3932,3,''),(258244,3932,4,''),(258245,3932,5,''),(258246,3932,6,''),(258247,3932,7,''),(258248,3932,8,''),(258249,3932,9,''),(258250,3932,10,''),(258251,3932,11,''),(258252,3932,12,''),(258253,3932,13,''),(258254,3932,14,''),(258255,3932,15,''),(258256,3932,16,''),(258257,3932,17,''),(258258,3932,18,''),(258259,3932,19,''),(258260,3932,20,''),(258261,3938,1,''),(258262,3938,2,''),(258263,3938,3,''),(258264,3938,4,''),(258265,3938,5,''),(258266,3938,6,''),(258267,3938,7,''),(258268,3938,8,''),(258269,3938,9,''),(258270,3938,10,''),(258271,3938,11,''),(258272,3938,12,''),(258273,3938,13,''),(258274,3938,14,''),(258275,3938,15,''),(258276,3938,16,''),(258277,3938,17,''),(258278,3938,18,''),(258279,3938,19,''),(258280,3938,20,''),(258281,3940,1,''),(258282,3940,2,''),(258283,3940,3,''),(258284,3940,4,''),(258285,3940,5,''),(258286,3940,6,''),(258287,3940,7,''),(258288,3940,8,''),(258289,3940,9,''),(258290,3940,10,''),(258291,3940,11,''),(258292,3940,12,''),(258293,3940,13,''),(258294,3940,14,''),(258295,3940,15,''),(258296,3940,16,''),(258297,3940,17,''),(258298,3940,18,''),(258299,3940,19,''),(258300,3940,20,''),(258301,3946,1,''),(258302,3946,2,''),(258303,3946,3,''),(258304,3946,4,''),(258305,3946,5,''),(258306,3946,6,''),(258307,3946,7,''),(258308,3946,8,''),(258309,3946,9,''),(258310,3946,10,''),(258311,3946,11,''),(258312,3946,12,''),(258313,3946,13,''),(258314,3946,14,''),(258315,3946,15,''),(258316,3946,16,''),(258317,3946,17,''),(258318,3946,18,''),(258319,3946,19,''),(258320,3946,20,''),(258321,3948,1,''),(258322,3948,2,''),(258323,3948,3,''),(258324,3948,4,''),(258325,3948,5,''),(258326,3948,6,''),(258327,3948,7,''),(258328,3948,8,''),(258329,3948,9,''),(258330,3948,10,''),(258331,3948,11,''),(258332,3948,12,''),(258333,3948,13,''),(258334,3948,14,''),(258335,3948,15,''),(258336,3948,16,''),(258337,3948,17,''),(258338,3948,18,''),(258339,3948,19,''),(258340,3948,20,''),(258341,3965,1,''),(258342,3965,2,''),(258343,3965,3,''),(258344,3965,4,''),(258345,3965,5,''),(258346,3965,6,''),(258347,3965,7,''),(258348,3965,8,''),(258349,3965,9,''),(258350,3965,10,''),(258351,3965,11,''),(258352,3965,12,''),(258353,3965,13,''),(258354,3965,14,''),(258355,3965,15,''),(258356,3965,16,''),(258357,3965,17,''),(258358,3965,18,''),(258359,3965,19,''),(258360,3965,20,''),(258361,3976,1,''),(258362,3976,2,''),(258363,3976,3,''),(258364,3976,4,''),(258365,3976,5,''),(258366,3976,6,''),(258367,3976,7,''),(258368,3976,8,''),(258369,3976,9,''),(258370,3976,10,''),(258371,3976,11,''),(258372,3976,12,''),(258373,3976,13,''),(258374,3976,14,''),(258375,3976,15,''),(258376,3976,16,''),(258377,3976,17,''),(258378,3976,18,''),(258379,3976,19,''),(258380,3976,20,''),(258381,3999,1,''),(258382,3999,2,''),(258383,3999,3,''),(258384,3999,4,''),(258385,3999,5,''),(258386,3999,6,''),(258387,3999,7,''),(258388,3999,8,''),(258389,3999,9,''),(258390,3999,10,''),(258391,3999,11,''),(258392,3999,12,''),(258393,3999,13,''),(258394,3999,14,''),(258395,3999,15,''),(258396,3999,16,''),(258397,3999,17,''),(258398,3999,18,''),(258399,3999,19,''),(258400,3999,20,''),(258401,4033,1,''),(258402,4033,2,''),(258403,4033,3,''),(258404,4033,4,''),(258405,4033,5,''),(258406,4033,6,''),(258407,4033,7,''),(258408,4033,8,''),(258409,4033,9,''),(258410,4033,10,''),(258411,4033,11,''),(258412,4033,12,''),(258413,4033,13,''),(258414,4033,14,''),(258415,4033,15,''),(258416,4033,16,''),(258417,4033,17,''),(258418,4033,18,''),(258419,4033,19,''),(258420,4033,20,''),(258421,4034,1,''),(258422,4034,2,''),(258423,4034,3,''),(258424,4034,4,''),(258425,4034,5,''),(258426,4034,6,''),(258427,4034,7,''),(258428,4034,8,''),(258429,4034,9,''),(258430,4034,10,''),(258431,4034,11,''),(258432,4034,12,''),(258433,4034,13,''),(258434,4034,14,''),(258435,4034,15,''),(258436,4034,16,''),(258437,4034,17,''),(258438,4034,18,''),(258439,4034,19,''),(258440,4034,20,''),(258441,4036,1,''),(258442,4036,2,''),(258443,4036,3,''),(258444,4036,4,''),(258445,4036,5,''),(258446,4036,6,''),(258447,4036,7,''),(258448,4036,8,''),(258449,4036,9,''),(258450,4036,10,''),(258451,4036,11,''),(258452,4036,12,''),(258453,4036,13,''),(258454,4036,14,''),(258455,4036,15,''),(258456,4036,16,''),(258457,4036,17,''),(258458,4036,18,''),(258459,4036,19,''),(258460,4036,20,''),(258461,4065,1,''),(258462,4065,2,''),(258463,4065,3,''),(258464,4065,4,''),(258465,4065,5,''),(258466,4065,6,''),(258467,4065,7,''),(258468,4065,8,''),(258469,4065,9,''),(258470,4065,10,''),(258471,4065,11,''),(258472,4065,12,''),(258473,4065,13,''),(258474,4065,14,''),(258475,4065,15,''),(258476,4065,16,''),(258477,4065,17,''),(258478,4065,18,''),(258479,4065,19,''),(258480,4065,20,''),(258481,4089,1,''),(258482,4089,2,''),(258483,4089,3,''),(258484,4089,4,''),(258485,4089,5,''),(258486,4089,6,''),(258487,4089,7,''),(258488,4089,8,''),(258489,4089,9,''),(258490,4089,10,''),(258491,4089,11,''),(258492,4089,12,''),(258493,4089,13,''),(258494,4089,14,''),(258495,4089,15,''),(258496,4089,16,''),(258497,4089,17,''),(258498,4089,18,''),(258499,4089,19,''),(258500,4089,20,''),(258501,4134,1,''),(258502,4134,2,''),(258503,4134,3,''),(258504,4134,4,''),(258505,4134,5,''),(258506,4134,6,''),(258507,4134,7,''),(258508,4134,8,''),(258509,4134,9,''),(258510,4134,10,''),(258511,4134,11,''),(258512,4134,12,''),(258513,4134,13,''),(258514,4134,14,''),(258515,4134,15,''),(258516,4134,16,''),(258517,4134,17,''),(258518,4134,18,''),(258519,4134,19,''),(258520,4134,20,''),(258521,4142,1,''),(258522,4142,2,''),(258523,4142,3,''),(258524,4142,4,''),(258525,4142,5,''),(258526,4142,6,''),(258527,4142,7,''),(258528,4142,8,''),(258529,4142,9,''),(258530,4142,10,''),(258531,4142,11,''),(258532,4142,12,''),(258533,4142,13,''),(258534,4142,14,''),(258535,4142,15,''),(258536,4142,16,''),(258537,4142,17,''),(258538,4142,18,''),(258539,4142,19,''),(258540,4142,20,''),(258541,4143,1,''),(258542,4143,2,''),(258543,4143,3,''),(258544,4143,4,''),(258545,4143,5,''),(258546,4143,6,''),(258547,4143,7,''),(258548,4143,8,''),(258549,4143,9,''),(258550,4143,10,''),(258551,4143,11,''),(258552,4143,12,''),(258553,4143,13,''),(258554,4143,14,''),(258555,4143,15,''),(258556,4143,16,''),(258557,4143,17,''),(258558,4143,18,''),(258559,4143,19,''),(258560,4143,20,''),(258561,4153,1,''),(258562,4153,2,''),(258563,4153,3,''),(258564,4153,4,''),(258565,4153,5,''),(258566,4153,6,''),(258567,4153,7,''),(258568,4153,8,''),(258569,4153,9,''),(258570,4153,10,''),(258571,4153,11,''),(258572,4153,12,''),(258573,4153,13,''),(258574,4153,14,''),(258575,4153,15,''),(258576,4153,16,''),(258577,4153,17,''),(258578,4153,18,''),(258579,4153,19,''),(258580,4153,20,''),(258581,4179,1,''),(258582,4179,2,''),(258583,4179,3,''),(258584,4179,4,''),(258585,4179,5,''),(258586,4179,6,''),(258587,4179,7,''),(258588,4179,8,''),(258589,4179,9,''),(258590,4179,10,''),(258591,4179,11,''),(258592,4179,12,''),(258593,4179,13,''),(258594,4179,14,''),(258595,4179,15,''),(258596,4179,16,''),(258597,4179,17,''),(258598,4179,18,''),(258599,4179,19,''),(258600,4179,20,''),(258601,4218,1,''),(258602,4218,2,''),(258603,4218,3,''),(258604,4218,4,''),(258605,4218,5,''),(258606,4218,6,''),(258607,4218,7,''),(258608,4218,8,''),(258609,4218,9,''),(258610,4218,10,''),(258611,4218,11,''),(258612,4218,12,''),(258613,4218,13,''),(258614,4218,14,''),(258615,4218,15,''),(258616,4218,16,''),(258617,4218,17,''),(258618,4218,18,''),(258619,4218,19,''),(258620,4218,20,''),(258621,4220,1,''),(258622,4220,2,''),(258623,4220,3,''),(258624,4220,4,''),(258625,4220,5,''),(258626,4220,6,''),(258627,4220,7,''),(258628,4220,8,''),(258629,4220,9,''),(258630,4220,10,''),(258631,4220,11,''),(258632,4220,12,''),(258633,4220,13,''),(258634,4220,14,''),(258635,4220,15,''),(258636,4220,16,''),(258637,4220,17,''),(258638,4220,18,''),(258639,4220,19,''),(258640,4220,20,''),(258641,4268,1,''),(258642,4268,2,''),(258643,4268,3,''),(258644,4268,4,''),(258645,4268,5,''),(258646,4268,6,''),(258647,4268,7,''),(258648,4268,8,''),(258649,4268,9,''),(258650,4268,10,''),(258651,4268,11,''),(258652,4268,12,''),(258653,4268,13,''),(258654,4268,14,''),(258655,4268,15,''),(258656,4268,16,''),(258657,4268,17,''),(258658,4268,18,''),(258659,4268,19,''),(258660,4268,20,''),(258661,4328,1,''),(258662,4328,2,''),(258663,4328,3,''),(258664,4328,4,''),(258665,4328,5,''),(258666,4328,6,''),(258667,4328,7,''),(258668,4328,8,''),(258669,4328,9,''),(258670,4328,10,''),(258671,4328,11,''),(258672,4328,12,''),(258673,4328,13,''),(258674,4328,14,''),(258675,4328,15,''),(258676,4328,16,''),(258677,4328,17,''),(258678,4328,18,''),(258679,4328,19,''),(258680,4328,20,''),(258681,4412,1,''),(258682,4412,2,''),(258683,4412,3,''),(258684,4412,4,''),(258685,4412,5,''),(258686,4412,6,''),(258687,4412,7,''),(258688,4412,8,''),(258689,4412,9,''),(258690,4412,10,''),(258691,4412,11,''),(258692,4412,12,''),(258693,4412,13,''),(258694,4412,14,''),(258695,4412,15,''),(258696,4412,16,''),(258697,4412,17,''),(258698,4412,18,''),(258699,4412,19,''),(258700,4412,20,''),(258701,4436,1,''),(258702,4436,2,''),(258703,4436,3,''),(258704,4436,4,''),(258705,4436,5,''),(258706,4436,6,''),(258707,4436,7,''),(258708,4436,8,''),(258709,4436,9,''),(258710,4436,10,''),(258711,4436,11,''),(258712,4436,12,''),(258713,4436,13,''),(258714,4436,14,''),(258715,4436,15,''),(258716,4436,16,''),(258717,4436,17,''),(258718,4436,18,''),(258719,4436,19,''),(258720,4436,20,''),(258721,4449,1,''),(258722,4449,2,''),(258723,4449,3,''),(258724,4449,4,''),(258725,4449,5,''),(258726,4449,6,''),(258727,4449,7,''),(258728,4449,8,''),(258729,4449,9,''),(258730,4449,10,''),(258731,4449,11,''),(258732,4449,12,''),(258733,4449,13,''),(258734,4449,14,''),(258735,4449,15,''),(258736,4449,16,''),(258737,4449,17,''),(258738,4449,18,''),(258739,4449,19,''),(258740,4449,20,''),(258741,4488,1,''),(258742,4488,2,''),(258743,4488,3,''),(258744,4488,4,''),(258745,4488,5,''),(258746,4488,6,''),(258747,4488,7,''),(258748,4488,8,''),(258749,4488,9,''),(258750,4488,10,''),(258751,4488,11,''),(258752,4488,12,''),(258753,4488,13,''),(258754,4488,14,''),(258755,4488,15,''),(258756,4488,16,''),(258757,4488,17,''),(258758,4488,18,''),(258759,4488,19,''),(258760,4488,20,''),(258761,4492,1,''),(258762,4492,2,''),(258763,4492,3,''),(258764,4492,4,''),(258765,4492,5,''),(258766,4492,6,''),(258767,4492,7,''),(258768,4492,8,''),(258769,4492,9,''),(258770,4492,10,''),(258771,4492,11,''),(258772,4492,12,''),(258773,4492,13,''),(258774,4492,14,''),(258775,4492,15,''),(258776,4492,16,''),(258777,4492,17,''),(258778,4492,18,''),(258779,4492,19,''),(258780,4492,20,''),(258781,4524,1,''),(258782,4524,2,''),(258783,4524,3,''),(258784,4524,4,''),(258785,4524,5,''),(258786,4524,6,''),(258787,4524,7,''),(258788,4524,8,''),(258789,4524,9,''),(258790,4524,10,''),(258791,4524,11,''),(258792,4524,12,''),(258793,4524,13,''),(258794,4524,14,''),(258795,4524,15,''),(258796,4524,16,''),(258797,4524,17,''),(258798,4524,18,''),(258799,4524,19,''),(258800,4524,20,''),(258801,4701,1,''),(258802,4701,2,''),(258803,4701,3,''),(258804,4701,4,''),(258805,4701,5,''),(258806,4701,6,''),(258807,4701,7,''),(258808,4701,8,''),(258809,4701,9,''),(258810,4701,10,''),(258811,4701,11,''),(258812,4701,12,''),(258813,4701,13,''),(258814,4701,14,''),(258815,4701,15,''),(258816,4701,16,''),(258817,4701,17,''),(258818,4701,18,''),(258819,4701,19,''),(258820,4701,20,''),(258821,4713,1,''),(258822,4713,2,''),(258823,4713,3,''),(258824,4713,4,''),(258825,4713,5,''),(258826,4713,6,''),(258827,4713,7,''),(258828,4713,8,''),(258829,4713,9,''),(258830,4713,10,''),(258831,4713,11,''),(258832,4713,12,''),(258833,4713,13,''),(258834,4713,14,''),(258835,4713,15,''),(258836,4713,16,''),(258837,4713,17,''),(258838,4713,18,''),(258839,4713,19,''),(258840,4713,20,''),(258841,4714,1,''),(258842,4714,2,''),(258843,4714,3,''),(258844,4714,4,''),(258845,4714,5,''),(258846,4714,6,''),(258847,4714,7,''),(258848,4714,8,''),(258849,4714,9,''),(258850,4714,10,''),(258851,4714,11,''),(258852,4714,12,''),(258853,4714,13,''),(258854,4714,14,''),(258855,4714,15,''),(258856,4714,16,''),(258857,4714,17,''),(258858,4714,18,''),(258859,4714,19,''),(258860,4714,20,''),(258861,4807,1,''),(258862,4807,2,''),(258863,4807,3,''),(258864,4807,4,''),(258865,4807,5,''),(258866,4807,6,''),(258867,4807,7,''),(258868,4807,8,''),(258869,4807,9,''),(258870,4807,10,''),(258871,4807,11,''),(258872,4807,12,''),(258873,4807,13,''),(258874,4807,14,''),(258875,4807,15,''),(258876,4807,16,''),(258877,4807,17,''),(258878,4807,18,''),(258879,4807,19,''),(258880,4807,20,''),(258881,4808,1,''),(258882,4808,2,''),(258883,4808,3,''),(258884,4808,4,''),(258885,4808,5,''),(258886,4808,6,''),(258887,4808,7,''),(258888,4808,8,''),(258889,4808,9,''),(258890,4808,10,''),(258891,4808,11,''),(258892,4808,12,''),(258893,4808,13,''),(258894,4808,14,''),(258895,4808,15,''),(258896,4808,16,''),(258897,4808,17,''),(258898,4808,18,''),(258899,4808,19,''),(258900,4808,20,''),(258901,4809,1,''),(258902,4809,2,''),(258903,4809,3,''),(258904,4809,4,''),(258905,4809,5,''),(258906,4809,6,''),(258907,4809,7,''),(258908,4809,8,''),(258909,4809,9,''),(258910,4809,10,''),(258911,4809,11,''),(258912,4809,12,''),(258913,4809,13,''),(258914,4809,14,''),(258915,4809,15,''),(258916,4809,16,''),(258917,4809,17,''),(258918,4809,18,''),(258919,4809,19,''),(258920,4809,20,''),(258921,4810,1,''),(258922,4810,2,''),(258923,4810,3,''),(258924,4810,4,''),(258925,4810,5,''),(258926,4810,6,''),(258927,4810,7,''),(258928,4810,8,''),(258929,4810,9,''),(258930,4810,10,''),(258931,4810,11,''),(258932,4810,12,''),(258933,4810,13,''),(258934,4810,14,''),(258935,4810,15,''),(258936,4810,16,''),(258937,4810,17,''),(258938,4810,18,''),(258939,4810,19,''),(258940,4810,20,''),(258941,4820,1,''),(258942,4820,2,''),(258943,4820,3,''),(258944,4820,4,''),(258945,4820,5,''),(258946,4820,6,''),(258947,4820,7,''),(258948,4820,8,''),(258949,4820,9,''),(258950,4820,10,''),(258951,4820,11,''),(258952,4820,12,''),(258953,4820,13,''),(258954,4820,14,''),(258955,4820,15,''),(258956,4820,16,''),(258957,4820,17,''),(258958,4820,18,''),(258959,4820,19,''),(258960,4820,20,''),(258961,4870,1,''),(258962,4870,2,''),(258963,4870,3,''),(258964,4870,4,''),(258965,4870,5,''),(258966,4870,6,''),(258967,4870,7,''),(258968,4870,8,''),(258969,4870,9,''),(258970,4870,10,''),(258971,4870,11,''),(258972,4870,12,''),(258973,4870,13,''),(258974,4870,14,''),(258975,4870,15,''),(258976,4870,16,''),(258977,4870,17,''),(258978,4870,18,''),(258979,4870,19,''),(258980,4870,20,''),(258981,4874,1,''),(258982,4874,2,''),(258983,4874,3,''),(258984,4874,4,''),(258985,4874,5,''),(258986,4874,6,''),(258987,4874,7,''),(258988,4874,8,''),(258989,4874,9,''),(258990,4874,10,''),(258991,4874,11,''),(258992,4874,12,''),(258993,4874,13,''),(258994,4874,14,''),(258995,4874,15,''),(258996,4874,16,''),(258997,4874,17,''),(258998,4874,18,''),(258999,4874,19,''),(259000,4874,20,''),(259001,4905,1,''),(259002,4905,2,''),(259003,4905,3,''),(259004,4905,4,''),(259005,4905,5,''),(259006,4905,6,''),(259007,4905,7,''),(259008,4905,8,''),(259009,4905,9,''),(259010,4905,10,''),(259011,4905,11,''),(259012,4905,12,''),(259013,4905,13,''),(259014,4905,14,''),(259015,4905,15,''),(259016,4905,16,''),(259017,4905,17,''),(259018,4905,18,''),(259019,4905,19,''),(259020,4905,20,''),(259021,4957,1,''),(259022,4957,2,''),(259023,4957,3,''),(259024,4957,4,''),(259025,4957,5,''),(259026,4957,6,''),(259027,4957,7,''),(259028,4957,8,''),(259029,4957,9,''),(259030,4957,10,''),(259031,4957,11,''),(259032,4957,12,''),(259033,4957,13,''),(259034,4957,14,''),(259035,4957,15,''),(259036,4957,16,''),(259037,4957,17,''),(259038,4957,18,''),(259039,4957,19,''),(259040,4957,20,''),(259041,5061,1,''),(259042,5061,2,''),(259043,5061,3,''),(259044,5061,4,''),(259045,5061,5,''),(259046,5061,6,''),(259047,5061,7,''),(259048,5061,8,''),(259049,5061,9,''),(259050,5061,10,''),(259051,5061,11,''),(259052,5061,12,''),(259053,5061,13,''),(259054,5061,14,''),(259055,5061,15,''),(259056,5061,16,''),(259057,5061,17,''),(259058,5061,18,''),(259059,5061,19,''),(259060,5061,20,''),(259061,5092,1,''),(259062,5092,2,''),(259063,5092,3,''),(259064,5092,4,''),(259065,5092,5,''),(259066,5092,6,''),(259067,5092,7,''),(259068,5092,8,''),(259069,5092,9,''),(259070,5092,10,''),(259071,5092,11,''),(259072,5092,12,''),(259073,5092,13,''),(259074,5092,14,''),(259075,5092,15,''),(259076,5092,16,''),(259077,5092,17,''),(259078,5092,18,''),(259079,5092,19,''),(259080,5092,20,''),(259081,5147,1,''),(259082,5147,2,''),(259083,5147,3,''),(259084,5147,4,''),(259085,5147,5,''),(259086,5147,6,''),(259087,5147,7,''),(259088,5147,8,''),(259089,5147,9,''),(259090,5147,10,''),(259091,5147,11,''),(259092,5147,12,''),(259093,5147,13,''),(259094,5147,14,''),(259095,5147,15,''),(259096,5147,16,''),(259097,5147,17,''),(259098,5147,18,''),(259099,5147,19,''),(259100,5147,20,''),(259101,5171,1,''),(259102,5171,2,''),(259103,5171,3,''),(259104,5171,4,''),(259105,5171,5,''),(259106,5171,6,''),(259107,5171,7,''),(259108,5171,8,''),(259109,5171,9,''),(259110,5171,10,''),(259111,5171,11,''),(259112,5171,12,''),(259113,5171,13,''),(259114,5171,14,''),(259115,5171,15,''),(259116,5171,16,''),(259117,5171,17,''),(259118,5171,18,''),(259119,5171,19,''),(259120,5171,20,''),(259121,5204,1,''),(259122,5204,2,''),(259123,5204,3,''),(259124,5204,4,''),(259125,5204,5,''),(259126,5204,6,''),(259127,5204,7,''),(259128,5204,8,''),(259129,5204,9,''),(259130,5204,10,''),(259131,5204,11,''),(259132,5204,12,''),(259133,5204,13,''),(259134,5204,14,''),(259135,5204,15,''),(259136,5204,16,''),(259137,5204,17,''),(259138,5204,18,''),(259139,5204,19,''),(259140,5204,20,''),(259141,5209,1,''),(259142,5209,2,''),(259143,5209,3,''),(259144,5209,4,''),(259145,5209,5,''),(259146,5209,6,''),(259147,5209,7,''),(259148,5209,8,''),(259149,5209,9,''),(259150,5209,10,''),(259151,5209,11,''),(259152,5209,12,''),(259153,5209,13,''),(259154,5209,14,''),(259155,5209,15,''),(259156,5209,16,''),(259157,5209,17,''),(259158,5209,18,''),(259159,5209,19,''),(259160,5209,20,''),(259161,5222,1,''),(259162,5222,2,''),(259163,5222,3,''),(259164,5222,4,''),(259165,5222,5,''),(259166,5222,6,''),(259167,5222,7,''),(259168,5222,8,''),(259169,5222,9,''),(259170,5222,10,''),(259171,5222,11,''),(259172,5222,12,''),(259173,5222,13,''),(259174,5222,14,''),(259175,5222,15,''),(259176,5222,16,''),(259177,5222,17,''),(259178,5222,18,''),(259179,5222,19,''),(259180,5222,20,''),(259181,5262,1,''),(259182,5262,2,''),(259183,5262,3,''),(259184,5262,4,''),(259185,5262,5,''),(259186,5262,6,''),(259187,5262,7,''),(259188,5262,8,''),(259189,5262,9,''),(259190,5262,10,''),(259191,5262,11,''),(259192,5262,12,''),(259193,5262,13,''),(259194,5262,14,''),(259195,5262,15,''),(259196,5262,16,''),(259197,5262,17,''),(259198,5262,18,''),(259199,5262,19,''),(259200,5262,20,''),(259201,5269,1,''),(259202,5269,2,''),(259203,5269,3,''),(259204,5269,4,''),(259205,5269,5,''),(259206,5269,6,''),(259207,5269,7,''),(259208,5269,8,''),(259209,5269,9,''),(259210,5269,10,''),(259211,5269,11,''),(259212,5269,12,''),(259213,5269,13,''),(259214,5269,14,''),(259215,5269,15,''),(259216,5269,16,''),(259217,5269,17,''),(259218,5269,18,''),(259219,5269,19,''),(259220,5269,20,''),(259221,5272,1,''),(259222,5272,2,''),(259223,5272,3,''),(259224,5272,4,''),(259225,5272,5,''),(259226,5272,6,''),(259227,5272,7,''),(259228,5272,8,''),(259229,5272,9,''),(259230,5272,10,''),(259231,5272,11,''),(259232,5272,12,''),(259233,5272,13,''),(259234,5272,14,''),(259235,5272,15,''),(259236,5272,16,''),(259237,5272,17,''),(259238,5272,18,''),(259239,5272,19,''),(259240,5272,20,''),(259241,5612,1,''),(259242,5612,2,''),(259243,5612,3,''),(259244,5612,4,''),(259245,5612,5,''),(259246,5612,6,''),(259247,5612,7,''),(259248,5612,8,''),(259249,5612,9,''),(259250,5612,10,''),(259251,5612,11,''),(259252,5612,12,''),(259253,5612,13,''),(259254,5612,14,''),(259255,5612,15,''),(259256,5612,16,''),(259257,5612,17,''),(259258,5612,18,''),(259259,5612,19,''),(259260,5612,20,''),(259261,5969,1,''),(259262,5969,2,''),(259263,5969,3,''),(259264,5969,4,''),(259265,5969,5,''),(259266,5969,6,''),(259267,5969,7,''),(259268,5969,8,''),(259269,5969,9,''),(259270,5969,10,''),(259271,5969,11,''),(259272,5969,12,''),(259273,5969,13,''),(259274,5969,14,''),(259275,5969,15,''),(259276,5969,16,''),(259277,5969,17,''),(259278,5969,18,''),(259279,5969,19,''),(259280,5969,20,''),(259281,6013,1,''),(259282,6013,2,''),(259283,6013,3,''),(259284,6013,4,''),(259285,6013,5,''),(259286,6013,6,''),(259287,6013,7,''),(259288,6013,8,''),(259289,6013,9,''),(259290,6013,10,''),(259291,6013,11,''),(259292,6013,12,''),(259293,6013,13,''),(259294,6013,14,''),(259295,6013,15,''),(259296,6013,16,''),(259297,6013,17,''),(259298,6013,18,''),(259299,6013,19,''),(259300,6013,20,''),(259301,6030,1,''),(259302,6030,2,''),(259303,6030,3,''),(259304,6030,4,''),(259305,6030,5,''),(259306,6030,6,''),(259307,6030,7,''),(259308,6030,8,''),(259309,6030,9,''),(259310,6030,10,''),(259311,6030,11,''),(259312,6030,12,''),(259313,6030,13,''),(259314,6030,14,''),(259315,6030,15,''),(259316,6030,16,''),(259317,6030,17,''),(259318,6030,18,''),(259319,6030,19,''),(259320,6030,20,''),(259321,6447,1,''),(259322,6447,2,''),(259323,6447,3,''),(259324,6447,4,''),(259325,6447,5,''),(259326,6447,6,''),(259327,6447,7,''),(259328,6447,8,''),(259329,6447,9,''),(259330,6447,10,''),(259331,6447,11,''),(259332,6447,12,''),(259333,6447,13,''),(259334,6447,14,''),(259335,6447,15,''),(259336,6447,16,''),(259337,6447,17,''),(259338,6447,18,''),(259339,6447,19,''),(259340,6447,20,''),(259341,6820,1,''),(259342,6820,2,''),(259343,6820,3,''),(259344,6820,4,''),(259345,6820,5,''),(259346,6820,6,''),(259347,6820,7,''),(259348,6820,8,''),(259349,6820,9,''),(259350,6820,10,''),(259351,6820,11,''),(259352,6820,12,''),(259353,6820,13,''),(259354,6820,14,''),(259355,6820,15,''),(259356,6820,16,''),(259357,6820,17,''),(259358,6820,18,''),(259359,6820,19,''),(259360,6820,20,''),(259361,7279,1,''),(259362,7279,2,''),(259363,7279,3,''),(259364,7279,4,''),(259365,7279,5,''),(259366,7279,6,''),(259367,7279,7,''),(259368,7279,8,''),(259369,7279,9,''),(259370,7279,10,''),(259371,7279,11,''),(259372,7279,12,''),(259373,7279,13,''),(259374,7279,14,''),(259375,7279,15,''),(259376,7279,16,''),(259377,7279,17,''),(259378,7279,18,''),(259379,7279,19,''),(259380,7279,20,''),(259381,7428,1,''),(259382,7428,2,''),(259383,7428,3,''),(259384,7428,4,''),(259385,7428,5,''),(259386,7428,6,''),(259387,7428,7,''),(259388,7428,8,''),(259389,7428,9,''),(259390,7428,10,''),(259391,7428,11,''),(259392,7428,12,''),(259393,7428,13,''),(259394,7428,14,''),(259395,7428,15,''),(259396,7428,16,''),(259397,7428,17,''),(259398,7428,18,''),(259399,7428,19,''),(259400,7428,20,''),(259401,7781,1,''),(259402,7781,2,''),(259403,7781,3,''),(259404,7781,4,''),(259405,7781,5,''),(259406,7781,6,''),(259407,7781,7,''),(259408,7781,8,''),(259409,7781,9,''),(259410,7781,10,''),(259411,7781,11,''),(259412,7781,12,''),(259413,7781,13,''),(259414,7781,14,''),(259415,7781,15,''),(259416,7781,16,''),(259417,7781,17,''),(259418,7781,18,''),(259419,7781,19,''),(259420,7781,20,''),(259421,8427,1,''),(259422,8427,2,''),(259423,8427,3,''),(259424,8427,4,''),(259425,8427,5,''),(259426,8427,6,''),(259427,8427,7,''),(259428,8427,8,''),(259429,8427,9,''),(259430,8427,10,''),(259431,8427,11,''),(259432,8427,12,''),(259433,8427,13,''),(259434,8427,14,''),(259435,8427,15,''),(259436,8427,16,''),(259437,8427,17,''),(259438,8427,18,''),(259439,8427,19,''),(259440,8427,20,''),(259441,9317,1,''),(259442,9317,2,''),(259443,9317,3,''),(259444,9317,4,''),(259445,9317,5,''),(259446,9317,6,''),(259447,9317,7,''),(259448,9317,8,''),(259449,9317,9,''),(259450,9317,10,''),(259451,9317,11,''),(259452,9317,12,''),(259453,9317,13,''),(259454,9317,14,''),(259455,9317,15,''),(259456,9317,16,''),(259457,9317,17,''),(259458,9317,18,''),(259459,9317,19,''),(259460,9317,20,''),(259461,9446,1,''),(259462,9446,2,''),(259463,9446,3,''),(259464,9446,4,''),(259465,9446,5,''),(259466,9446,6,''),(259467,9446,7,''),(259468,9446,8,''),(259469,9446,9,''),(259470,9446,10,''),(259471,9446,11,''),(259472,9446,12,''),(259473,9446,13,''),(259474,9446,14,''),(259475,9446,15,''),(259476,9446,16,''),(259477,9446,17,''),(259478,9446,18,''),(259479,9446,19,''),(259480,9446,20,''),(259481,9525,1,''),(259482,9525,2,''),(259483,9525,3,''),(259484,9525,4,''),(259485,9525,5,''),(259486,9525,6,''),(259487,9525,7,''),(259488,9525,8,''),(259489,9525,9,''),(259490,9525,10,''),(259491,9525,11,''),(259492,9525,12,''),(259493,9525,13,''),(259494,9525,14,''),(259495,9525,15,''),(259496,9525,16,''),(259497,9525,17,''),(259498,9525,18,''),(259499,9525,19,''),(259500,9525,20,''),(259501,9649,1,''),(259502,9649,2,''),(259503,9649,3,''),(259504,9649,4,''),(259505,9649,5,''),(259506,9649,6,''),(259507,9649,7,''),(259508,9649,8,''),(259509,9649,9,''),(259510,9649,10,''),(259511,9649,11,''),(259512,9649,12,''),(259513,9649,13,''),(259514,9649,14,''),(259515,9649,15,''),(259516,9649,16,''),(259517,9649,17,''),(259518,9649,18,''),(259519,9649,19,''),(259520,9649,20,''),(259521,9679,1,''),(259522,9679,2,''),(259523,9679,3,''),(259524,9679,4,''),(259525,9679,5,''),(259526,9679,6,''),(259527,9679,7,''),(259528,9679,8,''),(259529,9679,9,''),(259530,9679,10,''),(259531,9679,11,''),(259532,9679,12,''),(259533,9679,13,''),(259534,9679,14,''),(259535,9679,15,''),(259536,9679,16,''),(259537,9679,17,''),(259538,9679,18,''),(259539,9679,19,''),(259540,9679,20,''),(259541,9963,1,''),(259542,9963,2,''),(259543,9963,3,''),(259544,9963,4,''),(259545,9963,5,''),(259546,9963,6,''),(259547,9963,7,''),(259548,9963,8,''),(259549,9963,9,''),(259550,9963,10,''),(259551,9963,11,''),(259552,9963,12,''),(259553,9963,13,''),(259554,9963,14,''),(259555,9963,15,''),(259556,9963,16,''),(259557,9963,17,''),(259558,9963,18,''),(259559,9963,19,''),(259560,9963,20,''),(259561,9972,1,''),(259562,9972,2,''),(259563,9972,3,''),(259564,9972,4,''),(259565,9972,5,''),(259566,9972,6,''),(259567,9972,7,''),(259568,9972,8,''),(259569,9972,9,''),(259570,9972,10,''),(259571,9972,11,''),(259572,9972,12,''),(259573,9972,13,''),(259574,9972,14,''),(259575,9972,15,''),(259576,9972,16,''),(259577,9972,17,''),(259578,9972,18,''),(259579,9972,19,''),(259580,9972,20,''),(259581,9975,1,''),(259582,9975,2,''),(259583,9975,3,''),(259584,9975,4,''),(259585,9975,5,''),(259586,9975,6,''),(259587,9975,7,''),(259588,9975,8,''),(259589,9975,9,''),(259590,9975,10,''),(259591,9975,11,''),(259592,9975,12,''),(259593,9975,13,''),(259594,9975,14,''),(259595,9975,15,''),(259596,9975,16,''),(259597,9975,17,''),(259598,9975,18,''),(259599,9975,19,''),(259600,9975,20,''),(259601,9978,1,''),(259602,9978,2,''),(259603,9978,3,''),(259604,9978,4,''),(259605,9978,5,''),(259606,9978,6,''),(259607,9978,7,''),(259608,9978,8,''),(259609,9978,9,''),(259610,9978,10,''),(259611,9978,11,''),(259612,9978,12,''),(259613,9978,13,''),(259614,9978,14,''),(259615,9978,15,''),(259616,9978,16,''),(259617,9978,17,''),(259618,9978,18,''),(259619,9978,19,''),(259620,9978,20,''),(259621,10060,1,''),(259622,10060,2,''),(259623,10060,3,''),(259624,10060,4,''),(259625,10060,5,''),(259626,10060,6,''),(259627,10060,7,''),(259628,10060,8,''),(259629,10060,9,''),(259630,10060,10,''),(259631,10060,11,''),(259632,10060,12,''),(259633,10060,13,''),(259634,10060,14,''),(259635,10060,15,''),(259636,10060,16,''),(259637,10060,17,''),(259638,10060,18,''),(259639,10060,19,''),(259640,10060,20,''),(259641,10098,1,''),(259642,10098,2,''),(259643,10098,3,''),(259644,10098,4,''),(259645,10098,5,''),(259646,10098,6,''),(259647,10098,7,''),(259648,10098,8,''),(259649,10098,9,''),(259650,10098,10,''),(259651,10098,11,''),(259652,10098,12,''),(259653,10098,13,''),(259654,10098,14,''),(259655,10098,15,''),(259656,10098,16,''),(259657,10098,17,''),(259658,10098,18,''),(259659,10098,19,''),(259660,10098,20,''),(259661,10241,1,''),(259662,10241,2,''),(259663,10241,3,''),(259664,10241,4,''),(259665,10241,5,''),(259666,10241,6,''),(259667,10241,7,''),(259668,10241,8,''),(259669,10241,9,''),(259670,10241,10,''),(259671,10241,11,''),(259672,10241,12,''),(259673,10241,13,''),(259674,10241,14,''),(259675,10241,15,''),(259676,10241,16,''),(259677,10241,17,''),(259678,10241,18,''),(259679,10241,19,''),(259680,10241,20,''),(259681,10242,1,''),(259682,10242,2,''),(259683,10242,3,''),(259684,10242,4,''),(259685,10242,5,''),(259686,10242,6,''),(259687,10242,7,''),(259688,10242,8,''),(259689,10242,9,''),(259690,10242,10,''),(259691,10242,11,''),(259692,10242,12,''),(259693,10242,13,''),(259694,10242,14,''),(259695,10242,15,''),(259696,10242,16,''),(259697,10242,17,''),(259698,10242,18,''),(259699,10242,19,''),(259700,10242,20,''),(259701,10277,1,''),(259702,10277,2,''),(259703,10277,3,''),(259704,10277,4,''),(259705,10277,5,''),(259706,10277,6,''),(259707,10277,7,''),(259708,10277,8,''),(259709,10277,9,''),(259710,10277,10,''),(259711,10277,11,''),(259712,10277,12,''),(259713,10277,13,''),(259714,10277,14,''),(259715,10277,15,''),(259716,10277,16,''),(259717,10277,17,''),(259718,10277,18,''),(259719,10277,19,''),(259720,10277,20,''),(259721,10561,1,''),(259722,10561,2,''),(259723,10561,3,''),(259724,10561,4,''),(259725,10561,5,''),(259726,10561,6,''),(259727,10561,7,''),(259728,10561,8,''),(259729,10561,9,''),(259730,10561,10,''),(259731,10561,11,''),(259732,10561,12,''),(259733,10561,13,''),(259734,10561,14,''),(259735,10561,15,''),(259736,10561,16,''),(259737,10561,17,''),(259738,10561,18,''),(259739,10561,19,''),(259740,10561,20,''),(259741,10612,1,''),(259742,10612,2,''),(259743,10612,3,''),(259744,10612,4,''),(259745,10612,5,''),(259746,10612,6,''),(259747,10612,7,''),(259748,10612,8,''),(259749,10612,9,''),(259750,10612,10,''),(259751,10612,11,''),(259752,10612,12,''),(259753,10612,13,''),(259754,10612,14,''),(259755,10612,15,''),(259756,10612,16,''),(259757,10612,17,''),(259758,10612,18,''),(259759,10612,19,''),(259760,10612,20,''),(259761,10688,1,''),(259762,10688,2,''),(259763,10688,3,''),(259764,10688,4,''),(259765,10688,5,''),(259766,10688,6,''),(259767,10688,7,''),(259768,10688,8,''),(259769,10688,9,''),(259770,10688,10,''),(259771,10688,11,''),(259772,10688,12,''),(259773,10688,13,''),(259774,10688,14,''),(259775,10688,15,''),(259776,10688,16,''),(259777,10688,17,''),(259778,10688,18,''),(259779,10688,19,''),(259780,10688,20,''),(259781,10830,1,''),(259782,10830,2,''),(259783,10830,3,''),(259784,10830,4,''),(259785,10830,5,''),(259786,10830,6,''),(259787,10830,7,''),(259788,10830,8,''),(259789,10830,9,''),(259790,10830,10,''),(259791,10830,11,''),(259792,10830,12,''),(259793,10830,13,''),(259794,10830,14,''),(259795,10830,15,''),(259796,10830,16,''),(259797,10830,17,''),(259798,10830,18,''),(259799,10830,19,''),(259800,10830,20,''),(259801,11509,1,''),(259802,11509,2,''),(259803,11509,3,''),(259804,11509,4,''),(259805,11509,5,''),(259806,11509,6,''),(259807,11509,7,''),(259808,11509,8,''),(259809,11509,9,''),(259810,11509,10,''),(259811,11509,11,''),(259812,11509,12,''),(259813,11509,13,''),(259814,11509,14,''),(259815,11509,15,''),(259816,11509,16,''),(259817,11509,17,''),(259818,11509,18,''),(259819,11509,19,''),(259820,11509,20,''),(259821,11617,1,''),(259822,11617,2,''),(259823,11617,3,''),(259824,11617,4,''),(259825,11617,5,''),(259826,11617,6,''),(259827,11617,7,''),(259828,11617,8,''),(259829,11617,9,''),(259830,11617,10,''),(259831,11617,11,''),(259832,11617,12,''),(259833,11617,13,''),(259834,11617,14,''),(259835,11617,15,''),(259836,11617,16,''),(259837,11617,17,''),(259838,11617,18,''),(259839,11617,19,''),(259840,11617,20,''),(259841,11788,1,''),(259842,11788,2,''),(259843,11788,3,''),(259844,11788,4,''),(259845,11788,5,''),(259846,11788,6,''),(259847,11788,7,''),(259848,11788,8,''),(259849,11788,9,''),(259850,11788,10,''),(259851,11788,11,''),(259852,11788,12,''),(259853,11788,13,''),(259854,11788,14,''),(259855,11788,15,''),(259856,11788,16,''),(259857,11788,17,''),(259858,11788,18,''),(259859,11788,19,''),(259860,11788,20,''),(259861,11862,1,''),(259862,11862,2,''),(259863,11862,3,''),(259864,11862,4,''),(259865,11862,5,''),(259866,11862,6,''),(259867,11862,7,''),(259868,11862,8,''),(259869,11862,9,''),(259870,11862,10,''),(259871,11862,11,''),(259872,11862,12,''),(259873,11862,13,''),(259874,11862,14,''),(259875,11862,15,''),(259876,11862,16,''),(259877,11862,17,''),(259878,11862,18,''),(259879,11862,19,''),(259880,11862,20,''),(259881,11865,1,''),(259882,11865,2,''),(259883,11865,3,''),(259884,11865,4,''),(259885,11865,5,''),(259886,11865,6,''),(259887,11865,7,''),(259888,11865,8,''),(259889,11865,9,''),(259890,11865,10,''),(259891,11865,11,''),(259892,11865,12,''),(259893,11865,13,''),(259894,11865,14,''),(259895,11865,15,''),(259896,11865,16,''),(259897,11865,17,''),(259898,11865,18,''),(259899,11865,19,''),(259900,11865,20,''),(259901,12075,1,''),(259902,12075,2,''),(259903,12075,3,''),(259904,12075,4,''),(259905,12075,5,''),(259906,12075,6,''),(259907,12075,7,''),(259908,12075,8,''),(259909,12075,9,''),(259910,12075,10,''),(259911,12075,11,''),(259912,12075,12,''),(259913,12075,13,''),(259914,12075,14,''),(259915,12075,15,''),(259916,12075,16,''),(259917,12075,17,''),(259918,12075,18,''),(259919,12075,19,''),(259920,12075,20,''),(259921,12470,1,''),(259922,12470,2,''),(259923,12470,3,''),(259924,12470,4,''),(259925,12470,5,''),(259926,12470,6,''),(259927,12470,7,''),(259928,12470,8,''),(259929,12470,9,''),(259930,12470,10,''),(259931,12470,11,''),(259932,12470,12,''),(259933,12470,13,''),(259934,12470,14,''),(259935,12470,15,''),(259936,12470,16,''),(259937,12470,17,''),(259938,12470,18,''),(259939,12470,19,''),(259940,12470,20,''),(259941,12490,1,''),(259942,12490,2,''),(259943,12490,3,''),(259944,12490,4,''),(259945,12490,5,''),(259946,12490,6,''),(259947,12490,7,''),(259948,12490,8,''),(259949,12490,9,''),(259950,12490,10,''),(259951,12490,11,''),(259952,12490,12,''),(259953,12490,13,''),(259954,12490,14,''),(259955,12490,15,''),(259956,12490,16,''),(259957,12490,17,''),(259958,12490,18,''),(259959,12490,19,''),(259960,12490,20,''),(259961,12700,1,''),(259962,12700,2,''),(259963,12700,3,''),(259964,12700,4,''),(259965,12700,5,''),(259966,12700,6,''),(259967,12700,7,''),(259968,12700,8,''),(259969,12700,9,''),(259970,12700,10,''),(259971,12700,11,''),(259972,12700,12,''),(259973,12700,13,''),(259974,12700,14,''),(259975,12700,15,''),(259976,12700,16,''),(259977,12700,17,''),(259978,12700,18,''),(259979,12700,19,''),(259980,12700,20,''),(259981,12922,1,''),(259982,12922,2,''),(259983,12922,3,''),(259984,12922,4,''),(259985,12922,5,''),(259986,12922,6,''),(259987,12922,7,''),(259988,12922,8,''),(259989,12922,9,''),(259990,12922,10,''),(259991,12922,11,''),(259992,12922,12,''),(259993,12922,13,''),(259994,12922,14,''),(259995,12922,15,''),(259996,12922,16,''),(259997,12922,17,''),(259998,12922,18,''),(259999,12922,19,''),(260000,12922,20,''),(260001,13871,1,''),(260002,13871,2,''),(260003,13871,3,''),(260004,13871,4,''),(260005,13871,5,''),(260006,13871,6,''),(260007,13871,7,''),(260008,13871,8,''),(260009,13871,9,''),(260010,13871,10,''),(260011,13871,11,''),(260012,13871,12,''),(260013,13871,13,''),(260014,13871,14,''),(260015,13871,15,''),(260016,13871,16,''),(260017,13871,17,''),(260018,13871,18,''),(260019,13871,19,''),(260020,13871,20,''),(260021,13880,1,''),(260022,13880,2,''),(260023,13880,3,''),(260024,13880,4,''),(260025,13880,5,''),(260026,13880,6,''),(260027,13880,7,''),(260028,13880,8,''),(260029,13880,9,''),(260030,13880,10,''),(260031,13880,11,''),(260032,13880,12,''),(260033,13880,13,''),(260034,13880,14,''),(260035,13880,15,''),(260036,13880,16,''),(260037,13880,17,''),(260038,13880,18,''),(260039,13880,19,''),(260040,13880,20,''),(260041,13922,1,''),(260042,13922,2,''),(260043,13922,3,''),(260044,13922,4,''),(260045,13922,5,''),(260046,13922,6,''),(260047,13922,7,''),(260048,13922,8,''),(260049,13922,9,''),(260050,13922,10,''),(260051,13922,11,''),(260052,13922,12,''),(260053,13922,13,''),(260054,13922,14,''),(260055,13922,15,''),(260056,13922,16,''),(260057,13922,17,''),(260058,13922,18,''),(260059,13922,19,''),(260060,13922,20,''),(260061,13985,1,''),(260062,13985,2,''),(260063,13985,3,''),(260064,13985,4,''),(260065,13985,5,''),(260066,13985,6,''),(260067,13985,7,''),(260068,13985,8,''),(260069,13985,9,''),(260070,13985,10,''),(260071,13985,11,''),(260072,13985,12,''),(260073,13985,13,''),(260074,13985,14,''),(260075,13985,15,''),(260076,13985,16,''),(260077,13985,17,''),(260078,13985,18,''),(260079,13985,19,''),(260080,13985,20,''),(260081,14032,1,''),(260082,14032,2,''),(260083,14032,3,''),(260084,14032,4,''),(260085,14032,5,''),(260086,14032,6,''),(260087,14032,7,''),(260088,14032,8,''),(260089,14032,9,''),(260090,14032,10,''),(260091,14032,11,''),(260092,14032,12,''),(260093,14032,13,''),(260094,14032,14,''),(260095,14032,15,''),(260096,14032,16,''),(260097,14032,17,''),(260098,14032,18,''),(260099,14032,19,''),(260100,14032,20,''),(260101,14058,1,''),(260102,14058,2,''),(260103,14058,3,''),(260104,14058,4,''),(260105,14058,5,''),(260106,14058,6,''),(260107,14058,7,''),(260108,14058,8,''),(260109,14058,9,''),(260110,14058,10,''),(260111,14058,11,''),(260112,14058,12,''),(260113,14058,13,''),(260114,14058,14,''),(260115,14058,15,''),(260116,14058,16,''),(260117,14058,17,''),(260118,14058,18,''),(260119,14058,19,''),(260120,14058,20,''),(260121,14059,1,''),(260122,14059,2,''),(260123,14059,3,''),(260124,14059,4,''),(260125,14059,5,''),(260126,14059,6,''),(260127,14059,7,''),(260128,14059,8,''),(260129,14059,9,''),(260130,14059,10,''),(260131,14059,11,''),(260132,14059,12,''),(260133,14059,13,''),(260134,14059,14,''),(260135,14059,15,''),(260136,14059,16,''),(260137,14059,17,''),(260138,14059,18,''),(260139,14059,19,''),(260140,14059,20,''),(260141,14106,1,''),(260142,14106,2,''),(260143,14106,3,''),(260144,14106,4,''),(260145,14106,5,''),(260146,14106,6,''),(260147,14106,7,''),(260148,14106,8,''),(260149,14106,9,''),(260150,14106,10,''),(260151,14106,11,''),(260152,14106,12,''),(260153,14106,13,''),(260154,14106,14,''),(260155,14106,15,''),(260156,14106,16,''),(260157,14106,17,''),(260158,14106,18,''),(260159,14106,19,''),(260160,14106,20,''),(260161,14107,1,''),(260162,14107,2,''),(260163,14107,3,''),(260164,14107,4,''),(260165,14107,5,''),(260166,14107,6,''),(260167,14107,7,''),(260168,14107,8,''),(260169,14107,9,''),(260170,14107,10,''),(260171,14107,11,''),(260172,14107,12,''),(260173,14107,13,''),(260174,14107,14,''),(260175,14107,15,''),(260176,14107,16,''),(260177,14107,17,''),(260178,14107,18,''),(260179,14107,19,''),(260180,14107,20,''),(260181,14108,1,''),(260182,14108,2,''),(260183,14108,3,''),(260184,14108,4,''),(260185,14108,5,''),(260186,14108,6,''),(260187,14108,7,''),(260188,14108,8,''),(260189,14108,9,''),(260190,14108,10,''),(260191,14108,11,''),(260192,14108,12,''),(260193,14108,13,''),(260194,14108,14,''),(260195,14108,15,''),(260196,14108,16,''),(260197,14108,17,''),(260198,14108,18,''),(260199,14108,19,''),(260200,14108,20,''),(260201,14110,1,''),(260202,14110,2,''),(260203,14110,3,''),(260204,14110,4,''),(260205,14110,5,''),(260206,14110,6,''),(260207,14110,7,''),(260208,14110,8,''),(260209,14110,9,''),(260210,14110,10,''),(260211,14110,11,''),(260212,14110,12,''),(260213,14110,13,''),(260214,14110,14,''),(260215,14110,15,''),(260216,14110,16,''),(260217,14110,17,''),(260218,14110,18,''),(260219,14110,19,''),(260220,14110,20,''),(260221,14111,1,''),(260222,14111,2,''),(260223,14111,3,''),(260224,14111,4,''),(260225,14111,5,''),(260226,14111,6,''),(260227,14111,7,''),(260228,14111,8,''),(260229,14111,9,''),(260230,14111,10,''),(260231,14111,11,''),(260232,14111,12,''),(260233,14111,13,''),(260234,14111,14,''),(260235,14111,15,''),(260236,14111,16,''),(260237,14111,17,''),(260238,14111,18,''),(260239,14111,19,''),(260240,14111,20,''),(260241,14112,1,''),(260242,14112,2,''),(260243,14112,3,''),(260244,14112,4,''),(260245,14112,5,''),(260246,14112,6,''),(260247,14112,7,''),(260248,14112,8,''),(260249,14112,9,''),(260250,14112,10,''),(260251,14112,11,''),(260252,14112,12,''),(260253,14112,13,''),(260254,14112,14,''),(260255,14112,15,''),(260256,14112,16,''),(260257,14112,17,''),(260258,14112,18,''),(260259,14112,19,''),(260260,14112,20,''),(260261,14223,1,''),(260262,14223,2,''),(260263,14223,3,''),(260264,14223,4,''),(260265,14223,5,''),(260266,14223,6,''),(260267,14223,7,''),(260268,14223,8,''),(260269,14223,9,''),(260270,14223,10,''),(260271,14223,11,''),(260272,14223,12,''),(260273,14223,13,''),(260274,14223,14,''),(260275,14223,15,''),(260276,14223,16,''),(260277,14223,17,''),(260278,14223,18,''),(260279,14223,19,''),(260280,14223,20,''),(260281,14224,1,''),(260282,14224,2,''),(260283,14224,3,''),(260284,14224,4,''),(260285,14224,5,''),(260286,14224,6,''),(260287,14224,7,''),(260288,14224,8,''),(260289,14224,9,''),(260290,14224,10,''),(260291,14224,11,''),(260292,14224,12,''),(260293,14224,13,''),(260294,14224,14,''),(260295,14224,15,''),(260296,14224,16,''),(260297,14224,17,''),(260298,14224,18,''),(260299,14224,19,''),(260300,14224,20,''),(260301,14225,1,''),(260302,14225,2,''),(260303,14225,3,''),(260304,14225,4,''),(260305,14225,5,''),(260306,14225,6,''),(260307,14225,7,''),(260308,14225,8,''),(260309,14225,9,''),(260310,14225,10,''),(260311,14225,11,''),(260312,14225,12,''),(260313,14225,13,''),(260314,14225,14,''),(260315,14225,15,''),(260316,14225,16,''),(260317,14225,17,''),(260318,14225,18,''),(260319,14225,19,''),(260320,14225,20,''),(260321,14227,1,''),(260322,14227,2,''),(260323,14227,3,''),(260324,14227,4,''),(260325,14227,5,''),(260326,14227,6,''),(260327,14227,7,''),(260328,14227,8,''),(260329,14227,9,''),(260330,14227,10,''),(260331,14227,11,''),(260332,14227,12,''),(260333,14227,13,''),(260334,14227,14,''),(260335,14227,15,''),(260336,14227,16,''),(260337,14227,17,''),(260338,14227,18,''),(260339,14227,19,''),(260340,14227,20,''),(260341,14228,1,''),(260342,14228,2,''),(260343,14228,3,''),(260344,14228,4,''),(260345,14228,5,''),(260346,14228,6,''),(260347,14228,7,''),(260348,14228,8,''),(260349,14228,9,''),(260350,14228,10,''),(260351,14228,11,''),(260352,14228,12,''),(260353,14228,13,''),(260354,14228,14,''),(260355,14228,15,''),(260356,14228,16,''),(260357,14228,17,''),(260358,14228,18,''),(260359,14228,19,''),(260360,14228,20,''),(260361,14229,1,''),(260362,14229,2,''),(260363,14229,3,''),(260364,14229,4,''),(260365,14229,5,''),(260366,14229,6,''),(260367,14229,7,''),(260368,14229,8,''),(260369,14229,9,''),(260370,14229,10,''),(260371,14229,11,''),(260372,14229,12,''),(260373,14229,13,''),(260374,14229,14,''),(260375,14229,15,''),(260376,14229,16,''),(260377,14229,17,''),(260378,14229,18,''),(260379,14229,19,''),(260380,14229,20,''),(260381,15105,1,''),(260382,15105,2,''),(260383,15105,3,''),(260384,15105,4,''),(260385,15105,5,''),(260386,15105,6,''),(260387,15105,7,''),(260388,15105,8,''),(260389,15105,9,''),(260390,15105,10,''),(260391,15105,11,''),(260392,15105,12,''),(260393,15105,13,''),(260394,15105,14,''),(260395,15105,15,''),(260396,15105,16,''),(260397,15105,17,''),(260398,15105,18,''),(260399,15105,19,''),(260400,15105,20,''),(260401,15115,1,''),(260402,15115,2,''),(260403,15115,3,''),(260404,15115,4,''),(260405,15115,5,''),(260406,15115,6,''),(260407,15115,7,''),(260408,15115,8,''),(260409,15115,9,''),(260410,15115,10,''),(260411,15115,11,''),(260412,15115,12,''),(260413,15115,13,''),(260414,15115,14,''),(260415,15115,15,''),(260416,15115,16,''),(260417,15115,17,''),(260418,15115,18,''),(260419,15115,19,''),(260420,15115,20,''),(260421,15234,1,''),(260422,15234,2,''),(260423,15234,3,''),(260424,15234,4,''),(260425,15234,5,''),(260426,15234,6,''),(260427,15234,7,''),(260428,15234,8,''),(260429,15234,9,''),(260430,15234,10,''),(260431,15234,11,''),(260432,15234,12,''),(260433,15234,13,''),(260434,15234,14,''),(260435,15234,15,''),(260436,15234,16,''),(260437,15234,17,''),(260438,15234,18,''),(260439,15234,19,''),(260440,15234,20,''),(260441,15247,1,''),(260442,15247,2,''),(260443,15247,3,''),(260444,15247,4,''),(260445,15247,5,''),(260446,15247,6,''),(260447,15247,7,''),(260448,15247,8,''),(260449,15247,9,''),(260450,15247,10,''),(260451,15247,11,''),(260452,15247,12,''),(260453,15247,13,''),(260454,15247,14,''),(260455,15247,15,''),(260456,15247,16,''),(260457,15247,17,''),(260458,15247,18,''),(260459,15247,19,''),(260460,15247,20,''),(260461,15693,1,''),(260462,15693,2,''),(260463,15693,3,''),(260464,15693,4,''),(260465,15693,5,''),(260466,15693,6,''),(260467,15693,7,''),(260468,15693,8,''),(260469,15693,9,''),(260470,15693,10,''),(260471,15693,11,''),(260472,15693,12,''),(260473,15693,13,''),(260474,15693,14,''),(260475,15693,15,''),(260476,15693,16,''),(260477,15693,17,''),(260478,15693,18,''),(260479,15693,19,''),(260480,15693,20,''),(260481,15705,1,''),(260482,15705,2,''),(260483,15705,3,''),(260484,15705,4,''),(260485,15705,5,''),(260486,15705,6,''),(260487,15705,7,''),(260488,15705,8,''),(260489,15705,9,''),(260490,15705,10,''),(260491,15705,11,''),(260492,15705,12,''),(260493,15705,13,''),(260494,15705,14,''),(260495,15705,15,''),(260496,15705,16,''),(260497,15705,17,''),(260498,15705,18,''),(260499,15705,19,''),(260500,15705,20,''),(260501,15712,1,''),(260502,15712,2,''),(260503,15712,3,''),(260504,15712,4,''),(260505,15712,5,''),(260506,15712,6,''),(260507,15712,7,''),(260508,15712,8,''),(260509,15712,9,''),(260510,15712,10,''),(260511,15712,11,''),(260512,15712,12,''),(260513,15712,13,''),(260514,15712,14,''),(260515,15712,15,''),(260516,15712,16,''),(260517,15712,17,''),(260518,15712,18,''),(260519,15712,19,''),(260520,15712,20,''),(260521,15714,1,''),(260522,15714,2,''),(260523,15714,3,''),(260524,15714,4,''),(260525,15714,5,''),(260526,15714,6,''),(260527,15714,7,''),(260528,15714,8,''),(260529,15714,9,''),(260530,15714,10,''),(260531,15714,11,''),(260532,15714,12,''),(260533,15714,13,''),(260534,15714,14,''),(260535,15714,15,''),(260536,15714,16,''),(260537,15714,17,''),(260538,15714,18,''),(260539,15714,19,''),(260540,15714,20,''),(260541,15715,1,''),(260542,15715,2,''),(260543,15715,3,''),(260544,15715,4,''),(260545,15715,5,''),(260546,15715,6,''),(260547,15715,7,''),(260548,15715,8,''),(260549,15715,9,''),(260550,15715,10,''),(260551,15715,11,''),(260552,15715,12,''),(260553,15715,13,''),(260554,15715,14,''),(260555,15715,15,''),(260556,15715,16,''),(260557,15715,17,''),(260558,15715,18,''),(260559,15715,19,''),(260560,15715,20,''),(260561,15719,1,''),(260562,15719,2,''),(260563,15719,3,''),(260564,15719,4,''),(260565,15719,5,''),(260566,15719,6,''),(260567,15719,7,''),(260568,15719,8,''),(260569,15719,9,''),(260570,15719,10,''),(260571,15719,11,''),(260572,15719,12,''),(260573,15719,13,''),(260574,15719,14,''),(260575,15719,15,''),(260576,15719,16,''),(260577,15719,17,''),(260578,15719,18,''),(260579,15719,19,''),(260580,15719,20,''),(260581,15720,1,''),(260582,15720,2,''),(260583,15720,3,''),(260584,15720,4,''),(260585,15720,5,''),(260586,15720,6,''),(260587,15720,7,''),(260588,15720,8,''),(260589,15720,9,''),(260590,15720,10,''),(260591,15720,11,''),(260592,15720,12,''),(260593,15720,13,''),(260594,15720,14,''),(260595,15720,15,''),(260596,15720,16,''),(260597,15720,17,''),(260598,15720,18,''),(260599,15720,19,''),(260600,15720,20,''),(260601,15722,1,''),(260602,15722,2,''),(260603,15722,3,''),(260604,15722,4,''),(260605,15722,5,''),(260606,15722,6,''),(260607,15722,7,''),(260608,15722,8,''),(260609,15722,9,''),(260610,15722,10,''),(260611,15722,11,''),(260612,15722,12,''),(260613,15722,13,''),(260614,15722,14,''),(260615,15722,15,''),(260616,15722,16,''),(260617,15722,17,''),(260618,15722,18,''),(260619,15722,19,''),(260620,15722,20,''),(260621,15728,1,''),(260622,15728,2,''),(260623,15728,3,''),(260624,15728,4,''),(260625,15728,5,''),(260626,15728,6,''),(260627,15728,7,''),(260628,15728,8,''),(260629,15728,9,''),(260630,15728,10,''),(260631,15728,11,''),(260632,15728,12,''),(260633,15728,13,''),(260634,15728,14,''),(260635,15728,15,''),(260636,15728,16,''),(260637,15728,17,''),(260638,15728,18,''),(260639,15728,19,''),(260640,15728,20,''),(260641,15729,1,''),(260642,15729,2,''),(260643,15729,3,''),(260644,15729,4,''),(260645,15729,5,''),(260646,15729,6,''),(260647,15729,7,''),(260648,15729,8,''),(260649,15729,9,''),(260650,15729,10,''),(260651,15729,11,''),(260652,15729,12,''),(260653,15729,13,''),(260654,15729,14,''),(260655,15729,15,''),(260656,15729,16,''),(260657,15729,17,''),(260658,15729,18,''),(260659,15729,19,''),(260660,15729,20,''),(260661,15731,1,''),(260662,15731,2,''),(260663,15731,3,''),(260664,15731,4,''),(260665,15731,5,''),(260666,15731,6,''),(260667,15731,7,''),(260668,15731,8,''),(260669,15731,9,''),(260670,15731,10,''),(260671,15731,11,''),(260672,15731,12,''),(260673,15731,13,''),(260674,15731,14,''),(260675,15731,15,''),(260676,15731,16,''),(260677,15731,17,''),(260678,15731,18,''),(260679,15731,19,''),(260680,15731,20,''),(260681,15733,1,''),(260682,15733,2,''),(260683,15733,3,''),(260684,15733,4,''),(260685,15733,5,''),(260686,15733,6,''),(260687,15733,7,''),(260688,15733,8,''),(260689,15733,9,''),(260690,15733,10,''),(260691,15733,11,''),(260692,15733,12,''),(260693,15733,13,''),(260694,15733,14,''),(260695,15733,15,''),(260696,15733,16,''),(260697,15733,17,''),(260698,15733,18,''),(260699,15733,19,''),(260700,15733,20,''),(260701,15743,1,''),(260702,15743,2,''),(260703,15743,3,''),(260704,15743,4,''),(260705,15743,5,''),(260706,15743,6,''),(260707,15743,7,''),(260708,15743,8,''),(260709,15743,9,''),(260710,15743,10,''),(260711,15743,11,''),(260712,15743,12,''),(260713,15743,13,''),(260714,15743,14,''),(260715,15743,15,''),(260716,15743,16,''),(260717,15743,17,''),(260718,15743,18,''),(260719,15743,19,''),(260720,15743,20,''),(260721,15745,1,''),(260722,15745,2,''),(260723,15745,3,''),(260724,15745,4,''),(260725,15745,5,''),(260726,15745,6,''),(260727,15745,7,''),(260728,15745,8,''),(260729,15745,9,''),(260730,15745,10,''),(260731,15745,11,''),(260732,15745,12,''),(260733,15745,13,''),(260734,15745,14,''),(260735,15745,15,''),(260736,15745,16,''),(260737,15745,17,''),(260738,15745,18,''),(260739,15745,19,''),(260740,15745,20,''),(260741,15747,1,''),(260742,15747,2,''),(260743,15747,3,''),(260744,15747,4,''),(260745,15747,5,''),(260746,15747,6,''),(260747,15747,7,''),(260748,15747,8,''),(260749,15747,9,''),(260750,15747,10,''),(260751,15747,11,''),(260752,15747,12,''),(260753,15747,13,''),(260754,15747,14,''),(260755,15747,15,''),(260756,15747,16,''),(260757,15747,17,''),(260758,15747,18,''),(260759,15747,19,''),(260760,15747,20,''),(260761,15753,1,''),(260762,15753,2,''),(260763,15753,3,''),(260764,15753,4,''),(260765,15753,5,''),(260766,15753,6,''),(260767,15753,7,''),(260768,15753,8,''),(260769,15753,9,''),(260770,15753,10,''),(260771,15753,11,''),(260772,15753,12,''),(260773,15753,13,''),(260774,15753,14,''),(260775,15753,15,''),(260776,15753,16,''),(260777,15753,17,''),(260778,15753,18,''),(260779,15753,19,''),(260780,15753,20,''),(260781,15776,1,''),(260782,15776,2,''),(260783,15776,3,''),(260784,15776,4,''),(260785,15776,5,''),(260786,15776,6,''),(260787,15776,7,''),(260788,15776,8,''),(260789,15776,9,''),(260790,15776,10,''),(260791,15776,11,''),(260792,15776,12,''),(260793,15776,13,''),(260794,15776,14,''),(260795,15776,15,''),(260796,15776,16,''),(260797,15776,17,''),(260798,15776,18,''),(260799,15776,19,''),(260800,15776,20,''),(260801,16160,1,''),(260802,16160,2,''),(260803,16160,3,''),(260804,16160,4,''),(260805,16160,5,''),(260806,16160,6,''),(260807,16160,7,''),(260808,16160,8,''),(260809,16160,9,''),(260810,16160,10,''),(260811,16160,11,''),(260812,16160,12,''),(260813,16160,13,''),(260814,16160,14,''),(260815,16160,15,''),(260816,16160,16,''),(260817,16160,17,''),(260818,16160,18,''),(260819,16160,19,''),(260820,16160,20,''),(260821,16628,1,''),(260822,16628,2,''),(260823,16628,3,''),(260824,16628,4,''),(260825,16628,5,''),(260826,16628,6,''),(260827,16628,7,''),(260828,16628,8,''),(260829,16628,9,''),(260830,16628,10,''),(260831,16628,11,''),(260832,16628,12,''),(260833,16628,13,''),(260834,16628,14,''),(260835,16628,15,''),(260836,16628,16,''),(260837,16628,17,''),(260838,16628,18,''),(260839,16628,19,''),(260840,16628,20,''),(260841,17107,1,''),(260842,17107,2,''),(260843,17107,3,''),(260844,17107,4,''),(260845,17107,5,''),(260846,17107,6,''),(260847,17107,7,''),(260848,17107,8,''),(260849,17107,9,''),(260850,17107,10,''),(260851,17107,11,''),(260852,17107,12,''),(260853,17107,13,''),(260854,17107,14,''),(260855,17107,15,''),(260856,17107,16,''),(260857,17107,17,''),(260858,17107,18,''),(260859,17107,19,''),(260860,17107,20,''),(260861,17232,1,''),(260862,17232,2,''),(260863,17232,3,''),(260864,17232,4,''),(260865,17232,5,''),(260866,17232,6,''),(260867,17232,7,''),(260868,17232,8,''),(260869,17232,9,''),(260870,17232,10,''),(260871,17232,11,''),(260872,17232,12,''),(260873,17232,13,''),(260874,17232,14,''),(260875,17232,15,''),(260876,17232,16,''),(260877,17232,17,''),(260878,17232,18,''),(260879,17232,19,''),(260880,17232,20,''),(260881,17262,1,''),(260882,17262,2,''),(260883,17262,3,''),(260884,17262,4,''),(260885,17262,5,''),(260886,17262,6,''),(260887,17262,7,''),(260888,17262,8,''),(260889,17262,9,''),(260890,17262,10,''),(260891,17262,11,''),(260892,17262,12,''),(260893,17262,13,''),(260894,17262,14,''),(260895,17262,15,''),(260896,17262,16,''),(260897,17262,17,''),(260898,17262,18,''),(260899,17262,19,''),(260900,17262,20,''),(260901,17418,1,''),(260902,17418,2,''),(260903,17418,3,''),(260904,17418,4,''),(260905,17418,5,''),(260906,17418,6,''),(260907,17418,7,''),(260908,17418,8,''),(260909,17418,9,''),(260910,17418,10,''),(260911,17418,11,''),(260912,17418,12,''),(260913,17418,13,''),(260914,17418,14,''),(260915,17418,15,''),(260916,17418,16,''),(260917,17418,17,''),(260918,17418,18,''),(260919,17418,19,''),(260920,17418,20,''),(260921,17725,1,''),(260922,17725,2,''),(260923,17725,3,''),(260924,17725,4,''),(260925,17725,5,''),(260926,17725,6,''),(260927,17725,7,''),(260928,17725,8,''),(260929,17725,9,''),(260930,17725,10,''),(260931,17725,11,''),(260932,17725,12,''),(260933,17725,13,''),(260934,17725,14,''),(260935,17725,15,''),(260936,17725,16,''),(260937,17725,17,''),(260938,17725,18,''),(260939,17725,19,''),(260940,17725,20,''),(260941,17863,1,''),(260942,17863,2,''),(260943,17863,3,''),(260944,17863,4,''),(260945,17863,5,''),(260946,17863,6,''),(260947,17863,7,''),(260948,17863,8,''),(260949,17863,9,''),(260950,17863,10,''),(260951,17863,11,''),(260952,17863,12,''),(260953,17863,13,''),(260954,17863,14,''),(260955,17863,15,''),(260956,17863,16,''),(260957,17863,17,''),(260958,17863,18,''),(260959,17863,19,''),(260960,17863,20,''),(260961,17864,1,''),(260962,17864,2,''),(260963,17864,3,''),(260964,17864,4,''),(260965,17864,5,''),(260966,17864,6,''),(260967,17864,7,''),(260968,17864,8,''),(260969,17864,9,''),(260970,17864,10,''),(260971,17864,11,''),(260972,17864,12,''),(260973,17864,13,''),(260974,17864,14,''),(260975,17864,15,''),(260976,17864,16,''),(260977,17864,17,''),(260978,17864,18,''),(260979,17864,19,''),(260980,17864,20,''),(260981,18308,1,''),(260982,18308,2,''),(260983,18308,3,''),(260984,18308,4,''),(260985,18308,5,''),(260986,18308,6,''),(260987,18308,7,''),(260988,18308,8,''),(260989,18308,9,''),(260990,18308,10,''),(260991,18308,11,''),(260992,18308,12,''),(260993,18308,13,''),(260994,18308,14,''),(260995,18308,15,''),(260996,18308,16,''),(260997,18308,17,''),(260998,18308,18,''),(260999,18308,19,''),(261000,18308,20,''),(261001,18310,1,''),(261002,18310,2,''),(261003,18310,3,''),(261004,18310,4,''),(261005,18310,5,''),(261006,18310,6,''),(261007,18310,7,''),(261008,18310,8,''),(261009,18310,9,''),(261010,18310,10,''),(261011,18310,11,''),(261012,18310,12,''),(261013,18310,13,''),(261014,18310,14,''),(261015,18310,15,''),(261016,18310,16,''),(261017,18310,17,''),(261018,18310,18,''),(261019,18310,19,''),(261020,18310,20,''),(261021,18402,1,''),(261022,18402,2,''),(261023,18402,3,''),(261024,18402,4,''),(261025,18402,5,''),(261026,18402,6,''),(261027,18402,7,''),(261028,18402,8,''),(261029,18402,9,''),(261030,18402,10,''),(261031,18402,11,''),(261032,18402,12,''),(261033,18402,13,''),(261034,18402,14,''),(261035,18402,15,''),(261036,18402,16,''),(261037,18402,17,''),(261038,18402,18,''),(261039,18402,19,''),(261040,18402,20,''),(261041,18404,1,''),(261042,18404,2,''),(261043,18404,3,''),(261044,18404,4,''),(261045,18404,5,''),(261046,18404,6,''),(261047,18404,7,''),(261048,18404,8,''),(261049,18404,9,''),(261050,18404,10,''),(261051,18404,11,''),(261052,18404,12,''),(261053,18404,13,''),(261054,18404,14,''),(261055,18404,15,''),(261056,18404,16,''),(261057,18404,17,''),(261058,18404,18,''),(261059,18404,19,''),(261060,18404,20,''),(261061,18407,1,''),(261062,18407,2,''),(261063,18407,3,''),(261064,18407,4,''),(261065,18407,5,''),(261066,18407,6,''),(261067,18407,7,''),(261068,18407,8,''),(261069,18407,9,''),(261070,18407,10,''),(261071,18407,11,''),(261072,18407,12,''),(261073,18407,13,''),(261074,18407,14,''),(261075,18407,15,''),(261076,18407,16,''),(261077,18407,17,''),(261078,18407,18,''),(261079,18407,19,''),(261080,18407,20,''),(261081,18463,1,''),(261082,18463,2,''),(261083,18463,3,''),(261084,18463,4,''),(261085,18463,5,''),(261086,18463,6,''),(261087,18463,7,''),(261088,18463,8,''),(261089,18463,9,''),(261090,18463,10,''),(261091,18463,11,''),(261092,18463,12,''),(261093,18463,13,''),(261094,18463,14,''),(261095,18463,15,''),(261096,18463,16,''),(261097,18463,17,''),(261098,18463,18,''),(261099,18463,19,''),(261100,18463,20,''),(261101,18595,1,''),(261102,18595,2,''),(261103,18595,3,''),(261104,18595,4,''),(261105,18595,5,''),(261106,18595,6,''),(261107,18595,7,''),(261108,18595,8,''),(261109,18595,9,''),(261110,18595,10,''),(261111,18595,11,''),(261112,18595,12,''),(261113,18595,13,''),(261114,18595,14,''),(261115,18595,15,''),(261116,18595,16,''),(261117,18595,17,''),(261118,18595,18,''),(261119,18595,19,''),(261120,18595,20,''),(261121,18730,1,''),(261122,18730,2,''),(261123,18730,3,''),(261124,18730,4,''),(261125,18730,5,''),(261126,18730,6,''),(261127,18730,7,''),(261128,18730,8,''),(261129,18730,9,''),(261130,18730,10,''),(261131,18730,11,''),(261132,18730,12,''),(261133,18730,13,''),(261134,18730,14,''),(261135,18730,15,''),(261136,18730,16,''),(261137,18730,17,''),(261138,18730,18,''),(261139,18730,19,''),(261140,18730,20,''),(261141,18969,1,''),(261142,18969,2,''),(261143,18969,3,''),(261144,18969,4,''),(261145,18969,5,''),(261146,18969,6,''),(261147,18969,7,''),(261148,18969,8,''),(261149,18969,9,''),(261150,18969,10,''),(261151,18969,11,''),(261152,18969,12,''),(261153,18969,13,''),(261154,18969,14,''),(261155,18969,15,''),(261156,18969,16,''),(261157,18969,17,''),(261158,18969,18,''),(261159,18969,19,''),(261160,18969,20,''),(261161,19060,1,''),(261162,19060,2,''),(261163,19060,3,''),(261164,19060,4,''),(261165,19060,5,''),(261166,19060,6,''),(261167,19060,7,''),(261168,19060,8,''),(261169,19060,9,''),(261170,19060,10,''),(261171,19060,11,''),(261172,19060,12,''),(261173,19060,13,''),(261174,19060,14,''),(261175,19060,15,''),(261176,19060,16,''),(261177,19060,17,''),(261178,19060,18,''),(261179,19060,19,''),(261180,19060,20,''),(261181,19077,1,''),(261182,19077,2,''),(261183,19077,3,''),(261184,19077,4,''),(261185,19077,5,''),(261186,19077,6,''),(261187,19077,7,''),(261188,19077,8,''),(261189,19077,9,''),(261190,19077,10,''),(261191,19077,11,''),(261192,19077,12,''),(261193,19077,13,''),(261194,19077,14,''),(261195,19077,15,''),(261196,19077,16,''),(261197,19077,17,''),(261198,19077,18,''),(261199,19077,19,''),(261200,19077,20,''),(261201,19421,1,''),(261202,19421,2,''),(261203,19421,3,''),(261204,19421,4,''),(261205,19421,5,''),(261206,19421,6,''),(261207,19421,7,''),(261208,19421,8,''),(261209,19421,9,''),(261210,19421,10,''),(261211,19421,11,''),(261212,19421,12,''),(261213,19421,13,''),(261214,19421,14,''),(261215,19421,15,''),(261216,19421,16,''),(261217,19421,17,''),(261218,19421,18,''),(261219,19421,19,''),(261220,19421,20,''),(261221,19717,1,''),(261222,19717,2,''),(261223,19717,3,''),(261224,19717,4,''),(261225,19717,5,''),(261226,19717,6,''),(261227,19717,7,''),(261228,19717,8,''),(261229,19717,9,''),(261230,19717,10,''),(261231,19717,11,''),(261232,19717,12,''),(261233,19717,13,''),(261234,19717,14,''),(261235,19717,15,''),(261236,19717,16,''),(261237,19717,17,''),(261238,19717,18,''),(261239,19717,19,''),(261240,19717,20,''),(261241,20115,1,''),(261242,20115,2,''),(261243,20115,3,''),(261244,20115,4,''),(261245,20115,5,''),(261246,20115,6,''),(261247,20115,7,''),(261248,20115,8,''),(261249,20115,9,''),(261250,20115,10,''),(261251,20115,11,''),(261252,20115,12,''),(261253,20115,13,''),(261254,20115,14,''),(261255,20115,15,''),(261256,20115,16,''),(261257,20115,17,''),(261258,20115,18,''),(261259,20115,19,''),(261260,20115,20,''),(261261,20120,1,''),(261262,20120,2,''),(261263,20120,3,''),(261264,20120,4,''),(261265,20120,5,''),(261266,20120,6,''),(261267,20120,7,''),(261268,20120,8,''),(261269,20120,9,''),(261270,20120,10,''),(261271,20120,11,''),(261272,20120,12,''),(261273,20120,13,''),(261274,20120,14,''),(261275,20120,15,''),(261276,20120,16,''),(261277,20120,17,''),(261278,20120,18,''),(261279,20120,19,''),(261280,20120,20,''),(261281,20125,1,''),(261282,20125,2,''),(261283,20125,3,''),(261284,20125,4,''),(261285,20125,5,''),(261286,20125,6,''),(261287,20125,7,''),(261288,20125,8,''),(261289,20125,9,''),(261290,20125,10,''),(261291,20125,11,''),(261292,20125,12,''),(261293,20125,13,''),(261294,20125,14,''),(261295,20125,15,''),(261296,20125,16,''),(261297,20125,17,''),(261298,20125,18,''),(261299,20125,19,''),(261300,20125,20,''),(261301,20134,1,''),(261302,20134,2,''),(261303,20134,3,''),(261304,20134,4,''),(261305,20134,5,''),(261306,20134,6,''),(261307,20134,7,''),(261308,20134,8,''),(261309,20134,9,''),(261310,20134,10,''),(261311,20134,11,''),(261312,20134,12,''),(261313,20134,13,''),(261314,20134,14,''),(261315,20134,15,''),(261316,20134,16,''),(261317,20134,17,''),(261318,20134,18,''),(261319,20134,19,''),(261320,20134,20,''),(261321,20289,1,''),(261322,20289,2,''),(261323,20289,3,''),(261324,20289,4,''),(261325,20289,5,''),(261326,20289,6,''),(261327,20289,7,''),(261328,20289,8,''),(261329,20289,9,''),(261330,20289,10,''),(261331,20289,11,''),(261332,20289,12,''),(261333,20289,13,''),(261334,20289,14,''),(261335,20289,15,''),(261336,20289,16,''),(261337,20289,17,''),(261338,20289,18,''),(261339,20289,19,''),(261340,20289,20,''),(261341,21506,1,''),(261342,21506,2,''),(261343,21506,3,''),(261344,21506,4,''),(261345,21506,5,''),(261346,21506,6,''),(261347,21506,7,''),(261348,21506,8,''),(261349,21506,9,''),(261350,21506,10,''),(261351,21506,11,''),(261352,21506,12,''),(261353,21506,13,''),(261354,21506,14,''),(261355,21506,15,''),(261356,21506,16,''),(261357,21506,17,''),(261358,21506,18,''),(261359,21506,19,''),(261360,21506,20,''),(261361,21508,1,''),(261362,21508,2,''),(261363,21508,3,''),(261364,21508,4,''),(261365,21508,5,''),(261366,21508,6,''),(261367,21508,7,''),(261368,21508,8,''),(261369,21508,9,''),(261370,21508,10,''),(261371,21508,11,''),(261372,21508,12,''),(261373,21508,13,''),(261374,21508,14,''),(261375,21508,15,''),(261376,21508,16,''),(261377,21508,17,''),(261378,21508,18,''),(261379,21508,19,''),(261380,21508,20,''),(261381,21510,1,''),(261382,21510,2,''),(261383,21510,3,''),(261384,21510,4,''),(261385,21510,5,''),(261386,21510,6,''),(261387,21510,7,''),(261388,21510,8,''),(261389,21510,9,''),(261390,21510,10,''),(261391,21510,11,''),(261392,21510,12,''),(261393,21510,13,''),(261394,21510,14,''),(261395,21510,15,''),(261396,21510,16,''),(261397,21510,17,''),(261398,21510,18,''),(261399,21510,19,''),(261400,21510,20,''),(261401,21511,1,''),(261402,21511,2,''),(261403,21511,3,''),(261404,21511,4,''),(261405,21511,5,''),(261406,21511,6,''),(261407,21511,7,''),(261408,21511,8,''),(261409,21511,9,''),(261410,21511,10,''),(261411,21511,11,''),(261412,21511,12,''),(261413,21511,13,''),(261414,21511,14,''),(261415,21511,15,''),(261416,21511,16,''),(261417,21511,17,''),(261418,21511,18,''),(261419,21511,19,''),(261420,21511,20,''),(261421,21514,1,''),(261422,21514,2,''),(261423,21514,3,''),(261424,21514,4,''),(261425,21514,5,''),(261426,21514,6,''),(261427,21514,7,''),(261428,21514,8,''),(261429,21514,9,''),(261430,21514,10,''),(261431,21514,11,''),(261432,21514,12,''),(261433,21514,13,''),(261434,21514,14,''),(261435,21514,15,''),(261436,21514,16,''),(261437,21514,17,''),(261438,21514,18,''),(261439,21514,19,''),(261440,21514,20,''),(261441,21518,1,''),(261442,21518,2,''),(261443,21518,3,''),(261444,21518,4,''),(261445,21518,5,''),(261446,21518,6,''),(261447,21518,7,''),(261448,21518,8,''),(261449,21518,9,''),(261450,21518,10,''),(261451,21518,11,''),(261452,21518,12,''),(261453,21518,13,''),(261454,21518,14,''),(261455,21518,15,''),(261456,21518,16,''),(261457,21518,17,''),(261458,21518,18,''),(261459,21518,19,''),(261460,21518,20,''),(261461,21585,1,''),(261462,21585,2,''),(261463,21585,3,''),(261464,21585,4,''),(261465,21585,5,''),(261466,21585,6,''),(261467,21585,7,''),(261468,21585,8,''),(261469,21585,9,''),(261470,21585,10,''),(261471,21585,11,''),(261472,21585,12,''),(261473,21585,13,''),(261474,21585,14,''),(261475,21585,15,''),(261476,21585,16,''),(261477,21585,17,''),(261478,21585,18,''),(261479,21585,19,''),(261480,21585,20,''),(261481,21593,1,''),(261482,21593,2,''),(261483,21593,3,''),(261484,21593,4,''),(261485,21593,5,''),(261486,21593,6,''),(261487,21593,7,''),(261488,21593,8,''),(261489,21593,9,''),(261490,21593,10,''),(261491,21593,11,''),(261492,21593,12,''),(261493,21593,13,''),(261494,21593,14,''),(261495,21593,15,''),(261496,21593,16,''),(261497,21593,17,''),(261498,21593,18,''),(261499,21593,19,''),(261500,21593,20,''),(261501,23020,1,''),(261502,23020,2,''),(261503,23020,3,''),(261504,23020,4,''),(261505,23020,5,''),(261506,23020,6,''),(261507,23020,7,''),(261508,23020,8,''),(261509,23020,9,''),(261510,23020,10,''),(261511,23020,11,''),(261512,23020,12,''),(261513,23020,13,''),(261514,23020,14,''),(261515,23020,15,''),(261516,23020,16,''),(261517,23020,17,''),(261518,23020,18,''),(261519,23020,19,''),(261520,23020,20,''),(261521,23214,1,''),(261522,23214,2,''),(261523,23214,3,''),(261524,23214,4,''),(261525,23214,5,''),(261526,23214,6,''),(261527,23214,7,''),(261528,23214,8,''),(261529,23214,9,''),(261530,23214,10,''),(261531,23214,11,''),(261532,23214,12,''),(261533,23214,13,''),(261534,23214,14,''),(261535,23214,15,''),(261536,23214,16,''),(261537,23214,17,''),(261538,23214,18,''),(261539,23214,19,''),(261540,23214,20,''),(261541,23301,1,''),(261542,23301,2,''),(261543,23301,3,''),(261544,23301,4,''),(261545,23301,5,''),(261546,23301,6,''),(261547,23301,7,''),(261548,23301,8,''),(261549,23301,9,''),(261550,23301,10,''),(261551,23301,11,''),(261552,23301,12,''),(261553,23301,13,''),(261554,23301,14,''),(261555,23301,15,''),(261556,23301,16,''),(261557,23301,17,''),(261558,23301,18,''),(261559,23301,19,''),(261560,23301,20,''),(261561,23395,1,''),(261562,23395,2,''),(261563,23395,3,''),(261564,23395,4,''),(261565,23395,5,''),(261566,23395,6,''),(261567,23395,7,''),(261568,23395,8,''),(261569,23395,9,''),(261570,23395,10,''),(261571,23395,11,''),(261572,23395,12,''),(261573,23395,13,''),(261574,23395,14,''),(261575,23395,15,''),(261576,23395,16,''),(261577,23395,17,''),(261578,23395,18,''),(261579,23395,19,''),(261580,23395,20,''),(261581,23414,1,''),(261582,23414,2,''),(261583,23414,3,''),(261584,23414,4,''),(261585,23414,5,''),(261586,23414,6,''),(261587,23414,7,''),(261588,23414,8,''),(261589,23414,9,''),(261590,23414,10,''),(261591,23414,11,''),(261592,23414,12,''),(261593,23414,13,''),(261594,23414,14,''),(261595,23414,15,''),(261596,23414,16,''),(261597,23414,17,''),(261598,23414,18,''),(261599,23414,19,''),(261600,23414,20,''),(261601,23423,1,''),(261602,23423,2,''),(261603,23423,3,''),(261604,23423,4,''),(261605,23423,5,''),(261606,23423,6,''),(261607,23423,7,''),(261608,23423,8,''),(261609,23423,9,''),(261610,23423,10,''),(261611,23423,11,''),(261612,23423,12,''),(261613,23423,13,''),(261614,23423,14,''),(261615,23423,15,''),(261616,23423,16,''),(261617,23423,17,''),(261618,23423,18,''),(261619,23423,19,''),(261620,23423,20,''),(261621,23588,1,''),(261622,23588,2,''),(261623,23588,3,''),(261624,23588,4,''),(261625,23588,5,''),(261626,23588,6,''),(261627,23588,7,''),(261628,23588,8,''),(261629,23588,9,''),(261630,23588,10,''),(261631,23588,11,''),(261632,23588,12,''),(261633,23588,13,''),(261634,23588,14,''),(261635,23588,15,''),(261636,23588,16,''),(261637,23588,17,''),(261638,23588,18,''),(261639,23588,19,''),(261640,23588,20,''),(261641,23731,1,''),(261642,23731,2,''),(261643,23731,3,''),(261644,23731,4,''),(261645,23731,5,''),(261646,23731,6,''),(261647,23731,7,''),(261648,23731,8,''),(261649,23731,9,''),(261650,23731,10,''),(261651,23731,11,''),(261652,23731,12,''),(261653,23731,13,''),(261654,23731,14,''),(261655,23731,15,''),(261656,23731,16,''),(261657,23731,17,''),(261658,23731,18,''),(261659,23731,19,''),(261660,23731,20,''),(261661,23795,1,''),(261662,23795,2,''),(261663,23795,3,''),(261664,23795,4,''),(261665,23795,5,''),(261666,23795,6,''),(261667,23795,7,''),(261668,23795,8,''),(261669,23795,9,''),(261670,23795,10,''),(261671,23795,11,''),(261672,23795,12,''),(261673,23795,13,''),(261674,23795,14,''),(261675,23795,15,''),(261676,23795,16,''),(261677,23795,17,''),(261678,23795,18,''),(261679,23795,19,''),(261680,23795,20,''),(261681,23955,1,''),(261682,23955,2,''),(261683,23955,3,''),(261684,23955,4,''),(261685,23955,5,''),(261686,23955,6,''),(261687,23955,7,''),(261688,23955,8,''),(261689,23955,9,''),(261690,23955,10,''),(261691,23955,11,''),(261692,23955,12,''),(261693,23955,13,''),(261694,23955,14,''),(261695,23955,15,''),(261696,23955,16,''),(261697,23955,17,''),(261698,23955,18,''),(261699,23955,19,''),(261700,23955,20,''),(261701,24047,1,''),(261702,24047,2,''),(261703,24047,3,''),(261704,24047,4,''),(261705,24047,5,''),(261706,24047,6,''),(261707,24047,7,''),(261708,24047,8,''),(261709,24047,9,''),(261710,24047,10,''),(261711,24047,11,''),(261712,24047,12,''),(261713,24047,13,''),(261714,24047,14,''),(261715,24047,15,''),(261716,24047,16,''),(261717,24047,17,''),(261718,24047,18,''),(261719,24047,19,''),(261720,24047,20,''),(261721,24794,1,''),(261722,24794,2,''),(261723,24794,3,''),(261724,24794,4,''),(261725,24794,5,''),(261726,24794,6,''),(261727,24794,7,''),(261728,24794,8,''),(261729,24794,9,''),(261730,24794,10,''),(261731,24794,11,''),(261732,24794,12,''),(261733,24794,13,''),(261734,24794,14,''),(261735,24794,15,''),(261736,24794,16,''),(261737,24794,17,''),(261738,24794,18,''),(261739,24794,19,''),(261740,24794,20,''),(261741,25003,1,''),(261742,25003,2,''),(261743,25003,3,''),(261744,25003,4,''),(261745,25003,5,''),(261746,25003,6,''),(261747,25003,7,''),(261748,25003,8,''),(261749,25003,9,''),(261750,25003,10,''),(261751,25003,11,''),(261752,25003,12,''),(261753,25003,13,''),(261754,25003,14,''),(261755,25003,15,''),(261756,25003,16,''),(261757,25003,17,''),(261758,25003,18,''),(261759,25003,19,''),(261760,25003,20,''),(261761,25354,1,''),(261762,25354,2,''),(261763,25354,3,''),(261764,25354,4,''),(261765,25354,5,''),(261766,25354,6,''),(261767,25354,7,''),(261768,25354,8,''),(261769,25354,9,''),(261770,25354,10,''),(261771,25354,11,''),(261772,25354,12,''),(261773,25354,13,''),(261774,25354,14,''),(261775,25354,15,''),(261776,25354,16,''),(261777,25354,17,''),(261778,25354,18,''),(261779,25354,19,''),(261780,25354,20,''),(261781,25699,1,''),(261782,25699,2,''),(261783,25699,3,''),(261784,25699,4,''),(261785,25699,5,''),(261786,25699,6,''),(261787,25699,7,''),(261788,25699,8,''),(261789,25699,9,''),(261790,25699,10,''),(261791,25699,11,''),(261792,25699,12,''),(261793,25699,13,''),(261794,25699,14,''),(261795,25699,15,''),(261796,25699,16,''),(261797,25699,17,''),(261798,25699,18,''),(261799,25699,19,''),(261800,25699,20,''),(261801,25990,1,''),(261802,25990,2,''),(261803,25990,3,''),(261804,25990,4,''),(261805,25990,5,''),(261806,25990,6,''),(261807,25990,7,''),(261808,25990,8,''),(261809,25990,9,''),(261810,25990,10,''),(261811,25990,11,''),(261812,25990,12,''),(261813,25990,13,''),(261814,25990,14,''),(261815,25990,15,''),(261816,25990,16,''),(261817,25990,17,''),(261818,25990,18,''),(261819,25990,19,''),(261820,25990,20,''),(261821,26110,1,''),(261822,26110,2,''),(261823,26110,3,''),(261824,26110,4,''),(261825,26110,5,''),(261826,26110,6,''),(261827,26110,7,''),(261828,26110,8,''),(261829,26110,9,''),(261830,26110,10,''),(261831,26110,11,''),(261832,26110,12,''),(261833,26110,13,''),(261834,26110,14,''),(261835,26110,15,''),(261836,26110,16,''),(261837,26110,17,''),(261838,26110,18,''),(261839,26110,19,''),(261840,26110,20,''),(261841,26112,1,''),(261842,26112,2,''),(261843,26112,3,''),(261844,26112,4,''),(261845,26112,5,''),(261846,26112,6,''),(261847,26112,7,''),(261848,26112,8,''),(261849,26112,9,''),(261850,26112,10,''),(261851,26112,11,''),(261852,26112,12,''),(261853,26112,13,''),(261854,26112,14,''),(261855,26112,15,''),(261856,26112,16,''),(261857,26112,17,''),(261858,26112,18,''),(261859,26112,19,''),(261860,26112,20,''),(261861,26113,1,''),(261862,26113,2,''),(261863,26113,3,''),(261864,26113,4,''),(261865,26113,5,''),(261866,26113,6,''),(261867,26113,7,''),(261868,26113,8,''),(261869,26113,9,''),(261870,26113,10,''),(261871,26113,11,''),(261872,26113,12,''),(261873,26113,13,''),(261874,26113,14,''),(261875,26113,15,''),(261876,26113,16,''),(261877,26113,17,''),(261878,26113,18,''),(261879,26113,19,''),(261880,26113,20,''),(261881,26115,1,''),(261882,26115,2,''),(261883,26115,3,''),(261884,26115,4,''),(261885,26115,5,''),(261886,26115,6,''),(261887,26115,7,''),(261888,26115,8,''),(261889,26115,9,''),(261890,26115,10,''),(261891,26115,11,''),(261892,26115,12,''),(261893,26115,13,''),(261894,26115,14,''),(261895,26115,15,''),(261896,26115,16,''),(261897,26115,17,''),(261898,26115,18,''),(261899,26115,19,''),(261900,26115,20,''),(261901,26448,1,''),(261902,26448,2,''),(261903,26448,3,''),(261904,26448,4,''),(261905,26448,5,''),(261906,26448,6,''),(261907,26448,7,''),(261908,26448,8,''),(261909,26448,9,''),(261910,26448,10,''),(261911,26448,11,''),(261912,26448,12,''),(261913,26448,13,''),(261914,26448,14,''),(261915,26448,15,''),(261916,26448,16,''),(261917,26448,17,''),(261918,26448,18,''),(261919,26448,19,''),(261920,26448,20,''),(261921,26466,1,''),(261922,26466,2,''),(261923,26466,3,''),(261924,26466,4,''),(261925,26466,5,''),(261926,26466,6,''),(261927,26466,7,''),(261928,26466,8,''),(261929,26466,9,''),(261930,26466,10,''),(261931,26466,11,''),(261932,26466,12,''),(261933,26466,13,''),(261934,26466,14,''),(261935,26466,15,''),(261936,26466,16,''),(261937,26466,17,''),(261938,26466,18,''),(261939,26466,19,''),(261940,26466,20,''),(261941,26492,1,''),(261942,26492,2,''),(261943,26492,3,''),(261944,26492,4,''),(261945,26492,5,''),(261946,26492,6,''),(261947,26492,7,''),(261948,26492,8,''),(261949,26492,9,''),(261950,26492,10,''),(261951,26492,11,''),(261952,26492,12,''),(261953,26492,13,''),(261954,26492,14,''),(261955,26492,15,''),(261956,26492,16,''),(261957,26492,17,''),(261958,26492,18,''),(261959,26492,19,''),(261960,26492,20,''),(261961,26724,1,''),(261962,26724,2,''),(261963,26724,3,''),(261964,26724,4,''),(261965,26724,5,''),(261966,26724,6,''),(261967,26724,7,''),(261968,26724,8,''),(261969,26724,9,''),(261970,26724,10,''),(261971,26724,11,''),(261972,26724,12,''),(261973,26724,13,''),(261974,26724,14,''),(261975,26724,15,''),(261976,26724,16,''),(261977,26724,17,''),(261978,26724,18,''),(261979,26724,19,''),(261980,26724,20,''),(261981,26735,1,''),(261982,26735,2,''),(261983,26735,3,''),(261984,26735,4,''),(261985,26735,5,''),(261986,26735,6,''),(261987,26735,7,''),(261988,26735,8,''),(261989,26735,9,''),(261990,26735,10,''),(261991,26735,11,''),(261992,26735,12,''),(261993,26735,13,''),(261994,26735,14,''),(261995,26735,15,''),(261996,26735,16,''),(261997,26735,17,''),(261998,26735,18,''),(261999,26735,19,''),(262000,26735,20,''),(262001,26742,1,''),(262002,26742,2,''),(262003,26742,3,''),(262004,26742,4,''),(262005,26742,5,''),(262006,26742,6,''),(262007,26742,7,''),(262008,26742,8,''),(262009,26742,9,''),(262010,26742,10,''),(262011,26742,11,''),(262012,26742,12,''),(262013,26742,13,''),(262014,26742,14,''),(262015,26742,15,''),(262016,26742,16,''),(262017,26742,17,''),(262018,26742,18,''),(262019,26742,19,''),(262020,26742,20,''),(262021,26744,1,''),(262022,26744,2,''),(262023,26744,3,''),(262024,26744,4,''),(262025,26744,5,''),(262026,26744,6,''),(262027,26744,7,''),(262028,26744,8,''),(262029,26744,9,''),(262030,26744,10,''),(262031,26744,11,''),(262032,26744,12,''),(262033,26744,13,''),(262034,26744,14,''),(262035,26744,15,''),(262036,26744,16,''),(262037,26744,17,''),(262038,26744,18,''),(262039,26744,19,''),(262040,26744,20,''),(262041,26768,1,''),(262042,26768,2,''),(262043,26768,3,''),(262044,26768,4,''),(262045,26768,5,''),(262046,26768,6,''),(262047,26768,7,''),(262048,26768,8,''),(262049,26768,9,''),(262050,26768,10,''),(262051,26768,11,''),(262052,26768,12,''),(262053,26768,13,''),(262054,26768,14,''),(262055,26768,15,''),(262056,26768,16,''),(262057,26768,17,''),(262058,26768,18,''),(262059,26768,19,''),(262060,26768,20,''),(262061,26777,1,''),(262062,26777,2,''),(262063,26777,3,''),(262064,26777,4,''),(262065,26777,5,''),(262066,26777,6,''),(262067,26777,7,''),(262068,26777,8,''),(262069,26777,9,''),(262070,26777,10,''),(262071,26777,11,''),(262072,26777,12,''),(262073,26777,13,''),(262074,26777,14,''),(262075,26777,15,''),(262076,26777,16,''),(262077,26777,17,''),(262078,26777,18,''),(262079,26777,19,''),(262080,26777,20,''),(262081,26778,1,''),(262082,26778,2,''),(262083,26778,3,''),(262084,26778,4,''),(262085,26778,5,''),(262086,26778,6,''),(262087,26778,7,''),(262088,26778,8,''),(262089,26778,9,''),(262090,26778,10,''),(262091,26778,11,''),(262092,26778,12,''),(262093,26778,13,''),(262094,26778,14,''),(262095,26778,15,''),(262096,26778,16,''),(262097,26778,17,''),(262098,26778,18,''),(262099,26778,19,''),(262100,26778,20,''),(262101,26782,1,''),(262102,26782,2,''),(262103,26782,3,''),(262104,26782,4,''),(262105,26782,5,''),(262106,26782,6,''),(262107,26782,7,''),(262108,26782,8,''),(262109,26782,9,''),(262110,26782,10,''),(262111,26782,11,''),(262112,26782,12,''),(262113,26782,13,''),(262114,26782,14,''),(262115,26782,15,''),(262116,26782,16,''),(262117,26782,17,''),(262118,26782,18,''),(262119,26782,19,''),(262120,26782,20,''),(262121,26795,1,''),(262122,26795,2,''),(262123,26795,3,''),(262124,26795,4,''),(262125,26795,5,''),(262126,26795,6,''),(262127,26795,7,''),(262128,26795,8,''),(262129,26795,9,''),(262130,26795,10,''),(262131,26795,11,''),(262132,26795,12,''),(262133,26795,13,''),(262134,26795,14,''),(262135,26795,15,''),(262136,26795,16,''),(262137,26795,17,''),(262138,26795,18,''),(262139,26795,19,''),(262140,26795,20,''),(262141,27188,1,''),(262142,27188,2,''),(262143,27188,3,''),(262144,27188,4,''),(262145,27188,5,''),(262146,27188,6,''),(262147,27188,7,''),(262148,27188,8,''),(262149,27188,9,''),(262150,27188,10,''),(262151,27188,11,''),(262152,27188,12,''),(262153,27188,13,''),(262154,27188,14,''),(262155,27188,15,''),(262156,27188,16,''),(262157,27188,17,''),(262158,27188,18,''),(262159,27188,19,''),(262160,27188,20,''),(262161,27217,1,''),(262162,27217,2,''),(262163,27217,3,''),(262164,27217,4,''),(262165,27217,5,''),(262166,27217,6,''),(262167,27217,7,''),(262168,27217,8,''),(262169,27217,9,''),(262170,27217,10,''),(262171,27217,11,''),(262172,27217,12,''),(262173,27217,13,''),(262174,27217,14,''),(262175,27217,15,''),(262176,27217,16,''),(262177,27217,17,''),(262178,27217,18,''),(262179,27217,19,''),(262180,27217,20,''),(262181,27218,1,''),(262182,27218,2,''),(262183,27218,3,''),(262184,27218,4,''),(262185,27218,5,''),(262186,27218,6,''),(262187,27218,7,''),(262188,27218,8,''),(262189,27218,9,''),(262190,27218,10,''),(262191,27218,11,''),(262192,27218,12,''),(262193,27218,13,''),(262194,27218,14,''),(262195,27218,15,''),(262196,27218,16,''),(262197,27218,17,''),(262198,27218,18,''),(262199,27218,19,''),(262200,27218,20,''),(262201,27307,1,''),(262202,27307,2,''),(262203,27307,3,''),(262204,27307,4,''),(262205,27307,5,''),(262206,27307,6,''),(262207,27307,7,''),(262208,27307,8,''),(262209,27307,9,''),(262210,27307,10,''),(262211,27307,11,''),(262212,27307,12,''),(262213,27307,13,''),(262214,27307,14,''),(262215,27307,15,''),(262216,27307,16,''),(262217,27307,17,''),(262218,27307,18,''),(262219,27307,19,''),(262220,27307,20,''),(262221,27340,1,''),(262222,27340,2,''),(262223,27340,3,''),(262224,27340,4,''),(262225,27340,5,''),(262226,27340,6,''),(262227,27340,7,''),(262228,27340,8,''),(262229,27340,9,''),(262230,27340,10,''),(262231,27340,11,''),(262232,27340,12,''),(262233,27340,13,''),(262234,27340,14,''),(262235,27340,15,''),(262236,27340,16,''),(262237,27340,17,''),(262238,27340,18,''),(262239,27340,19,''),(262240,27340,20,''),(262241,27786,1,''),(262242,27786,2,''),(262243,27786,3,''),(262244,27786,4,''),(262245,27786,5,''),(262246,27786,6,''),(262247,27786,7,''),(262248,27786,8,''),(262249,27786,9,''),(262250,27786,10,''),(262251,27786,11,''),(262252,27786,12,''),(262253,27786,13,''),(262254,27786,14,''),(262255,27786,15,''),(262256,27786,16,''),(262257,27786,17,''),(262258,27786,18,''),(262259,27786,19,''),(262260,27786,20,''),(262261,27787,1,''),(262262,27787,2,''),(262263,27787,3,''),(262264,27787,4,''),(262265,27787,5,''),(262266,27787,6,''),(262267,27787,7,''),(262268,27787,8,''),(262269,27787,9,''),(262270,27787,10,''),(262271,27787,11,''),(262272,27787,12,''),(262273,27787,13,''),(262274,27787,14,''),(262275,27787,15,''),(262276,27787,16,''),(262277,27787,17,''),(262278,27787,18,''),(262279,27787,19,''),(262280,27787,20,''),(262281,27788,1,''),(262282,27788,2,''),(262283,27788,3,''),(262284,27788,4,''),(262285,27788,5,''),(262286,27788,6,''),(262287,27788,7,''),(262288,27788,8,''),(262289,27788,9,''),(262290,27788,10,''),(262291,27788,11,''),(262292,27788,12,''),(262293,27788,13,''),(262294,27788,14,''),(262295,27788,15,''),(262296,27788,16,''),(262297,27788,17,''),(262298,27788,18,''),(262299,27788,19,''),(262300,27788,20,''),(262301,27789,1,''),(262302,27789,2,''),(262303,27789,3,''),(262304,27789,4,''),(262305,27789,5,''),(262306,27789,6,''),(262307,27789,7,''),(262308,27789,8,''),(262309,27789,9,''),(262310,27789,10,''),(262311,27789,11,''),(262312,27789,12,''),(262313,27789,13,''),(262314,27789,14,''),(262315,27789,15,''),(262316,27789,16,''),(262317,27789,17,''),(262318,27789,18,''),(262319,27789,19,''),(262320,27789,20,''),(262321,28571,1,''),(262322,28571,2,''),(262323,28571,3,''),(262324,28571,4,''),(262325,28571,5,''),(262326,28571,6,''),(262327,28571,7,''),(262328,28571,8,''),(262329,28571,9,''),(262330,28571,10,''),(262331,28571,11,''),(262332,28571,12,''),(262333,28571,13,''),(262334,28571,14,''),(262335,28571,15,''),(262336,28571,16,''),(262337,28571,17,''),(262338,28571,18,''),(262339,28571,19,''),(262340,28571,20,''),(262341,28573,1,''),(262342,28573,2,''),(262343,28573,3,''),(262344,28573,4,''),(262345,28573,5,''),(262346,28573,6,''),(262347,28573,7,''),(262348,28573,8,''),(262349,28573,9,''),(262350,28573,10,''),(262351,28573,11,''),(262352,28573,12,''),(262353,28573,13,''),(262354,28573,14,''),(262355,28573,15,''),(262356,28573,16,''),(262357,28573,17,''),(262358,28573,18,''),(262359,28573,19,''),(262360,28573,20,''),(262361,28820,1,''),(262362,28820,2,''),(262363,28820,3,''),(262364,28820,4,''),(262365,28820,5,''),(262366,28820,6,''),(262367,28820,7,''),(262368,28820,8,''),(262369,28820,9,''),(262370,28820,10,''),(262371,28820,11,''),(262372,28820,12,''),(262373,28820,13,''),(262374,28820,14,''),(262375,28820,15,''),(262376,28820,16,''),(262377,28820,17,''),(262378,28820,18,''),(262379,28820,19,''),(262380,28820,20,''),(262381,28821,1,''),(262382,28821,2,''),(262383,28821,3,''),(262384,28821,4,''),(262385,28821,5,''),(262386,28821,6,''),(262387,28821,7,''),(262388,28821,8,''),(262389,28821,9,''),(262390,28821,10,''),(262391,28821,11,''),(262392,28821,12,''),(262393,28821,13,''),(262394,28821,14,''),(262395,28821,15,''),(262396,28821,16,''),(262397,28821,17,''),(262398,28821,18,''),(262399,28821,19,''),(262400,28821,20,''),(262401,28824,1,''),(262402,28824,2,''),(262403,28824,3,''),(262404,28824,4,''),(262405,28824,5,''),(262406,28824,6,''),(262407,28824,7,''),(262408,28824,8,''),(262409,28824,9,''),(262410,28824,10,''),(262411,28824,11,''),(262412,28824,12,''),(262413,28824,13,''),(262414,28824,14,''),(262415,28824,15,''),(262416,28824,16,''),(262417,28824,17,''),(262418,28824,18,''),(262419,28824,19,''),(262420,28824,20,''),(262421,28825,1,''),(262422,28825,2,''),(262423,28825,3,''),(262424,28825,4,''),(262425,28825,5,''),(262426,28825,6,''),(262427,28825,7,''),(262428,28825,8,''),(262429,28825,9,''),(262430,28825,10,''),(262431,28825,11,''),(262432,28825,12,''),(262433,28825,13,''),(262434,28825,14,''),(262435,28825,15,''),(262436,28825,16,''),(262437,28825,17,''),(262438,28825,18,''),(262439,28825,19,''),(262440,28825,20,''),(262441,28827,1,''),(262442,28827,2,''),(262443,28827,3,''),(262444,28827,4,''),(262445,28827,5,''),(262446,28827,6,''),(262447,28827,7,''),(262448,28827,8,''),(262449,28827,9,''),(262450,28827,10,''),(262451,28827,11,''),(262452,28827,12,''),(262453,28827,13,''),(262454,28827,14,''),(262455,28827,15,''),(262456,28827,16,''),(262457,28827,17,''),(262458,28827,18,''),(262459,28827,19,''),(262460,28827,20,''),(262461,28828,1,''),(262462,28828,2,''),(262463,28828,3,''),(262464,28828,4,''),(262465,28828,5,''),(262466,28828,6,''),(262467,28828,7,''),(262468,28828,8,''),(262469,28828,9,''),(262470,28828,10,''),(262471,28828,11,''),(262472,28828,12,''),(262473,28828,13,''),(262474,28828,14,''),(262475,28828,15,''),(262476,28828,16,''),(262477,28828,17,''),(262478,28828,18,''),(262479,28828,19,''),(262480,28828,20,''),(262481,28830,1,''),(262482,28830,2,''),(262483,28830,3,''),(262484,28830,4,''),(262485,28830,5,''),(262486,28830,6,''),(262487,28830,7,''),(262488,28830,8,''),(262489,28830,9,''),(262490,28830,10,''),(262491,28830,11,''),(262492,28830,12,''),(262493,28830,13,''),(262494,28830,14,''),(262495,28830,15,''),(262496,28830,16,''),(262497,28830,17,''),(262498,28830,18,''),(262499,28830,19,''),(262500,28830,20,''),(262501,28831,1,''),(262502,28831,2,''),(262503,28831,3,''),(262504,28831,4,''),(262505,28831,5,''),(262506,28831,6,''),(262507,28831,7,''),(262508,28831,8,''),(262509,28831,9,''),(262510,28831,10,''),(262511,28831,11,''),(262512,28831,12,''),(262513,28831,13,''),(262514,28831,14,''),(262515,28831,15,''),(262516,28831,16,''),(262517,28831,17,''),(262518,28831,18,''),(262519,28831,19,''),(262520,28831,20,''),(262521,28832,1,''),(262522,28832,2,''),(262523,28832,3,''),(262524,28832,4,''),(262525,28832,5,''),(262526,28832,6,''),(262527,28832,7,''),(262528,28832,8,''),(262529,28832,9,''),(262530,28832,10,''),(262531,28832,11,''),(262532,28832,12,''),(262533,28832,13,''),(262534,28832,14,''),(262535,28832,15,''),(262536,28832,16,''),(262537,28832,17,''),(262538,28832,18,''),(262539,28832,19,''),(262540,28832,20,''),(262541,28834,1,''),(262542,28834,2,''),(262543,28834,3,''),(262544,28834,4,''),(262545,28834,5,''),(262546,28834,6,''),(262547,28834,7,''),(262548,28834,8,''),(262549,28834,9,''),(262550,28834,10,''),(262551,28834,11,''),(262552,28834,12,''),(262553,28834,13,''),(262554,28834,14,''),(262555,28834,15,''),(262556,28834,16,''),(262557,28834,17,''),(262558,28834,18,''),(262559,28834,19,''),(262560,28834,20,''),(262561,29323,1,''),(262562,29323,2,''),(262563,29323,3,''),(262564,29323,4,''),(262565,29323,5,''),(262566,29323,6,''),(262567,29323,7,''),(262568,29323,8,''),(262569,29323,9,''),(262570,29323,10,''),(262571,29323,11,''),(262572,29323,12,''),(262573,29323,13,''),(262574,29323,14,''),(262575,29323,15,''),(262576,29323,16,''),(262577,29323,17,''),(262578,29323,18,''),(262579,29323,19,''),(262580,29323,20,''),(262581,29391,1,''),(262582,29391,2,''),(262583,29391,3,''),(262584,29391,4,''),(262585,29391,5,''),(262586,29391,6,''),(262587,29391,7,''),(262588,29391,8,''),(262589,29391,9,''),(262590,29391,10,''),(262591,29391,11,''),(262592,29391,12,''),(262593,29391,13,''),(262594,29391,14,''),(262595,29391,15,''),(262596,29391,16,''),(262597,29391,17,''),(262598,29391,18,''),(262599,29391,19,''),(262600,29391,20,''),(262601,29405,1,''),(262602,29405,2,''),(262603,29405,3,''),(262604,29405,4,''),(262605,29405,5,''),(262606,29405,6,''),(262607,29405,7,''),(262608,29405,8,''),(262609,29405,9,''),(262610,29405,10,''),(262611,29405,11,''),(262612,29405,12,''),(262613,29405,13,''),(262614,29405,14,''),(262615,29405,15,''),(262616,29405,16,''),(262617,29405,17,''),(262618,29405,18,''),(262619,29405,19,''),(262620,29405,20,''),(262621,29431,1,''),(262622,29431,2,''),(262623,29431,3,''),(262624,29431,4,''),(262625,29431,5,''),(262626,29431,6,''),(262627,29431,7,''),(262628,29431,8,''),(262629,29431,9,''),(262630,29431,10,''),(262631,29431,11,''),(262632,29431,12,''),(262633,29431,13,''),(262634,29431,14,''),(262635,29431,15,''),(262636,29431,16,''),(262637,29431,17,''),(262638,29431,18,''),(262639,29431,19,''),(262640,29431,20,''),(262641,29455,1,''),(262642,29455,2,''),(262643,29455,3,''),(262644,29455,4,''),(262645,29455,5,''),(262646,29455,6,''),(262647,29455,7,''),(262648,29455,8,''),(262649,29455,9,''),(262650,29455,10,''),(262651,29455,11,''),(262652,29455,12,''),(262653,29455,13,''),(262654,29455,14,''),(262655,29455,15,''),(262656,29455,16,''),(262657,29455,17,''),(262658,29455,18,''),(262659,29455,19,''),(262660,29455,20,''),(262661,29961,1,''),(262662,29961,2,''),(262663,29961,3,''),(262664,29961,4,''),(262665,29961,5,''),(262666,29961,6,''),(262667,29961,7,''),(262668,29961,8,''),(262669,29961,9,''),(262670,29961,10,''),(262671,29961,11,''),(262672,29961,12,''),(262673,29961,13,''),(262674,29961,14,''),(262675,29961,15,''),(262676,29961,16,''),(262677,29961,17,''),(262678,29961,18,''),(262679,29961,19,''),(262680,29961,20,''),(262681,29969,1,''),(262682,29969,2,''),(262683,29969,3,''),(262684,29969,4,''),(262685,29969,5,''),(262686,29969,6,''),(262687,29969,7,''),(262688,29969,8,''),(262689,29969,9,''),(262690,29969,10,''),(262691,29969,11,''),(262692,29969,12,''),(262693,29969,13,''),(262694,29969,14,''),(262695,29969,15,''),(262696,29969,16,''),(262697,29969,17,''),(262698,29969,18,''),(262699,29969,19,''),(262700,29969,20,''),(262701,30096,1,''),(262702,30096,2,''),(262703,30096,3,''),(262704,30096,4,''),(262705,30096,5,''),(262706,30096,6,''),(262707,30096,7,''),(262708,30096,8,''),(262709,30096,9,''),(262710,30096,10,''),(262711,30096,11,''),(262712,30096,12,''),(262713,30096,13,''),(262714,30096,14,''),(262715,30096,15,''),(262716,30096,16,''),(262717,30096,17,''),(262718,30096,18,''),(262719,30096,19,''),(262720,30096,20,''),(262721,30287,1,''),(262722,30287,2,''),(262723,30287,3,''),(262724,30287,4,''),(262725,30287,5,''),(262726,30287,6,''),(262727,30287,7,''),(262728,30287,8,''),(262729,30287,9,''),(262730,30287,10,''),(262731,30287,11,''),(262732,30287,12,''),(262733,30287,13,''),(262734,30287,14,''),(262735,30287,15,''),(262736,30287,16,''),(262737,30287,17,''),(262738,30287,18,''),(262739,30287,19,''),(262740,30287,20,''),(262741,30291,1,''),(262742,30291,2,''),(262743,30291,3,''),(262744,30291,4,''),(262745,30291,5,''),(262746,30291,6,''),(262747,30291,7,''),(262748,30291,8,''),(262749,30291,9,''),(262750,30291,10,''),(262751,30291,11,''),(262752,30291,12,''),(262753,30291,13,''),(262754,30291,14,''),(262755,30291,15,''),(262756,30291,16,''),(262757,30291,17,''),(262758,30291,18,''),(262759,30291,19,''),(262760,30291,20,''),(262761,30292,1,''),(262762,30292,2,''),(262763,30292,3,''),(262764,30292,4,''),(262765,30292,5,''),(262766,30292,6,''),(262767,30292,7,''),(262768,30292,8,''),(262769,30292,9,''),(262770,30292,10,''),(262771,30292,11,''),(262772,30292,12,''),(262773,30292,13,''),(262774,30292,14,''),(262775,30292,15,''),(262776,30292,16,''),(262777,30292,17,''),(262778,30292,18,''),(262779,30292,19,''),(262780,30292,20,''),(262781,30293,1,''),(262782,30293,2,''),(262783,30293,3,''),(262784,30293,4,''),(262785,30293,5,''),(262786,30293,6,''),(262787,30293,7,''),(262788,30293,8,''),(262789,30293,9,''),(262790,30293,10,''),(262791,30293,11,''),(262792,30293,12,''),(262793,30293,13,''),(262794,30293,14,''),(262795,30293,15,''),(262796,30293,16,''),(262797,30293,17,''),(262798,30293,18,''),(262799,30293,19,''),(262800,30293,20,''),(262801,30294,1,''),(262802,30294,2,''),(262803,30294,3,''),(262804,30294,4,''),(262805,30294,5,''),(262806,30294,6,''),(262807,30294,7,''),(262808,30294,8,''),(262809,30294,9,''),(262810,30294,10,''),(262811,30294,11,''),(262812,30294,12,''),(262813,30294,13,''),(262814,30294,14,''),(262815,30294,15,''),(262816,30294,16,''),(262817,30294,17,''),(262818,30294,18,''),(262819,30294,19,''),(262820,30294,20,''),(262821,30335,1,''),(262822,30335,2,''),(262823,30335,3,''),(262824,30335,4,''),(262825,30335,5,''),(262826,30335,6,''),(262827,30335,7,''),(262828,30335,8,''),(262829,30335,9,''),(262830,30335,10,''),(262831,30335,11,''),(262832,30335,12,''),(262833,30335,13,''),(262834,30335,14,''),(262835,30335,15,''),(262836,30335,16,''),(262837,30335,17,''),(262838,30335,18,''),(262839,30335,19,''),(262840,30335,20,''),(262841,30662,1,''),(262842,30662,2,''),(262843,30662,3,''),(262844,30662,4,''),(262845,30662,5,''),(262846,30662,6,''),(262847,30662,7,''),(262848,30662,8,''),(262849,30662,9,''),(262850,30662,10,''),(262851,30662,11,''),(262852,30662,12,''),(262853,30662,13,''),(262854,30662,14,''),(262855,30662,15,''),(262856,30662,16,''),(262857,30662,17,''),(262858,30662,18,''),(262859,30662,19,''),(262860,30662,20,''),(262861,30665,1,''),(262862,30665,2,''),(262863,30665,3,''),(262864,30665,4,''),(262865,30665,5,''),(262866,30665,6,''),(262867,30665,7,''),(262868,30665,8,''),(262869,30665,9,''),(262870,30665,10,''),(262871,30665,11,''),(262872,30665,12,''),(262873,30665,13,''),(262874,30665,14,''),(262875,30665,15,''),(262876,30665,16,''),(262877,30665,17,''),(262878,30665,18,''),(262879,30665,19,''),(262880,30665,20,''),(262881,30712,1,''),(262882,30712,2,''),(262883,30712,3,''),(262884,30712,4,''),(262885,30712,5,''),(262886,30712,6,''),(262887,30712,7,''),(262888,30712,8,''),(262889,30712,9,''),(262890,30712,10,''),(262891,30712,11,''),(262892,30712,12,''),(262893,30712,13,''),(262894,30712,14,''),(262895,30712,15,''),(262896,30712,16,''),(262897,30712,17,''),(262898,30712,18,''),(262899,30712,19,''),(262900,30712,20,''),(262901,30760,1,''),(262902,30760,2,''),(262903,30760,3,''),(262904,30760,4,''),(262905,30760,5,''),(262906,30760,6,''),(262907,30760,7,''),(262908,30760,8,''),(262909,30760,9,''),(262910,30760,10,''),(262911,30760,11,''),(262912,30760,12,''),(262913,30760,13,''),(262914,30760,14,''),(262915,30760,15,''),(262916,30760,16,''),(262917,30760,17,''),(262918,30760,18,''),(262919,30760,19,''),(262920,30760,20,''),(262921,30863,1,''),(262922,30863,2,''),(262923,30863,3,''),(262924,30863,4,''),(262925,30863,5,''),(262926,30863,6,''),(262927,30863,7,''),(262928,30863,8,''),(262929,30863,9,''),(262930,30863,10,''),(262931,30863,11,''),(262932,30863,12,''),(262933,30863,13,''),(262934,30863,14,''),(262935,30863,15,''),(262936,30863,16,''),(262937,30863,17,''),(262938,30863,18,''),(262939,30863,19,''),(262940,30863,20,''),(262941,31200,1,''),(262942,31200,2,''),(262943,31200,3,''),(262944,31200,4,''),(262945,31200,5,''),(262946,31200,6,''),(262947,31200,7,''),(262948,31200,8,''),(262949,31200,9,''),(262950,31200,10,''),(262951,31200,11,''),(262952,31200,12,''),(262953,31200,13,''),(262954,31200,14,''),(262955,31200,15,''),(262956,31200,16,''),(262957,31200,17,''),(262958,31200,18,''),(262959,31200,19,''),(262960,31200,20,''),(262961,31237,1,''),(262962,31237,2,''),(262963,31237,3,''),(262964,31237,4,''),(262965,31237,5,''),(262966,31237,6,''),(262967,31237,7,''),(262968,31237,8,''),(262969,31237,9,''),(262970,31237,10,''),(262971,31237,11,''),(262972,31237,12,''),(262973,31237,13,''),(262974,31237,14,''),(262975,31237,15,''),(262976,31237,16,''),(262977,31237,17,''),(262978,31237,18,''),(262979,31237,19,''),(262980,31237,20,''),(262981,31241,1,''),(262982,31241,2,''),(262983,31241,3,''),(262984,31241,4,''),(262985,31241,5,''),(262986,31241,6,''),(262987,31241,7,''),(262988,31241,8,''),(262989,31241,9,''),(262990,31241,10,''),(262991,31241,11,''),(262992,31241,12,''),(262993,31241,13,''),(262994,31241,14,''),(262995,31241,15,''),(262996,31241,16,''),(262997,31241,17,''),(262998,31241,18,''),(262999,31241,19,''),(263000,31241,20,''),(263001,31269,1,''),(263002,31269,2,''),(263003,31269,3,''),(263004,31269,4,''),(263005,31269,5,''),(263006,31269,6,''),(263007,31269,7,''),(263008,31269,8,''),(263009,31269,9,''),(263010,31269,10,''),(263011,31269,11,''),(263012,31269,12,''),(263013,31269,13,''),(263014,31269,14,''),(263015,31269,15,''),(263016,31269,16,''),(263017,31269,17,''),(263018,31269,18,''),(263019,31269,19,''),(263020,31269,20,''),(263021,31692,1,''),(263022,31692,2,''),(263023,31692,3,''),(263024,31692,4,''),(263025,31692,5,''),(263026,31692,6,''),(263027,31692,7,''),(263028,31692,8,''),(263029,31692,9,''),(263030,31692,10,''),(263031,31692,11,''),(263032,31692,12,''),(263033,31692,13,''),(263034,31692,14,''),(263035,31692,15,''),(263036,31692,16,''),(263037,31692,17,''),(263038,31692,18,''),(263039,31692,19,''),(263040,31692,20,''),(263041,31693,1,''),(263042,31693,2,''),(263043,31693,3,''),(263044,31693,4,''),(263045,31693,5,''),(263046,31693,6,''),(263047,31693,7,''),(263048,31693,8,''),(263049,31693,9,''),(263050,31693,10,''),(263051,31693,11,''),(263052,31693,12,''),(263053,31693,13,''),(263054,31693,14,''),(263055,31693,15,''),(263056,31693,16,''),(263057,31693,17,''),(263058,31693,18,''),(263059,31693,19,''),(263060,31693,20,''),(263061,31696,1,''),(263062,31696,2,''),(263063,31696,3,''),(263064,31696,4,''),(263065,31696,5,''),(263066,31696,6,''),(263067,31696,7,''),(263068,31696,8,''),(263069,31696,9,''),(263070,31696,10,''),(263071,31696,11,''),(263072,31696,12,''),(263073,31696,13,''),(263074,31696,14,''),(263075,31696,15,''),(263076,31696,16,''),(263077,31696,17,''),(263078,31696,18,''),(263079,31696,19,''),(263080,31696,20,''),(263081,31749,1,''),(263082,31749,2,''),(263083,31749,3,''),(263084,31749,4,''),(263085,31749,5,''),(263086,31749,6,''),(263087,31749,7,''),(263088,31749,8,''),(263089,31749,9,''),(263090,31749,10,''),(263091,31749,11,''),(263092,31749,12,''),(263093,31749,13,''),(263094,31749,14,''),(263095,31749,15,''),(263096,31749,16,''),(263097,31749,17,''),(263098,31749,18,''),(263099,31749,19,''),(263100,31749,20,''),(263101,31814,1,''),(263102,31814,2,''),(263103,31814,3,''),(263104,31814,4,''),(263105,31814,5,''),(263106,31814,6,''),(263107,31814,7,''),(263108,31814,8,''),(263109,31814,9,''),(263110,31814,10,''),(263111,31814,11,''),(263112,31814,12,''),(263113,31814,13,''),(263114,31814,14,''),(263115,31814,15,''),(263116,31814,16,''),(263117,31814,17,''),(263118,31814,18,''),(263119,31814,19,''),(263120,31814,20,''),(263121,31815,1,''),(263122,31815,2,''),(263123,31815,3,''),(263124,31815,4,''),(263125,31815,5,''),(263126,31815,6,''),(263127,31815,7,''),(263128,31815,8,''),(263129,31815,9,''),(263130,31815,10,''),(263131,31815,11,''),(263132,31815,12,''),(263133,31815,13,''),(263134,31815,14,''),(263135,31815,15,''),(263136,31815,16,''),(263137,31815,17,''),(263138,31815,18,''),(263139,31815,19,''),(263140,31815,20,''),(263141,31820,1,''),(263142,31820,2,''),(263143,31820,3,''),(263144,31820,4,''),(263145,31820,5,''),(263146,31820,6,''),(263147,31820,7,''),(263148,31820,8,''),(263149,31820,9,''),(263150,31820,10,''),(263151,31820,11,''),(263152,31820,12,''),(263153,31820,13,''),(263154,31820,14,''),(263155,31820,15,''),(263156,31820,16,''),(263157,31820,17,''),(263158,31820,18,''),(263159,31820,19,''),(263160,31820,20,''),(263161,31821,1,''),(263162,31821,2,''),(263163,31821,3,''),(263164,31821,4,''),(263165,31821,5,''),(263166,31821,6,''),(263167,31821,7,''),(263168,31821,8,''),(263169,31821,9,''),(263170,31821,10,''),(263171,31821,11,''),(263172,31821,12,''),(263173,31821,13,''),(263174,31821,14,''),(263175,31821,15,''),(263176,31821,16,''),(263177,31821,17,''),(263178,31821,18,''),(263179,31821,19,''),(263180,31821,20,''),(263181,31823,1,''),(263182,31823,2,''),(263183,31823,3,''),(263184,31823,4,''),(263185,31823,5,''),(263186,31823,6,''),(263187,31823,7,''),(263188,31823,8,''),(263189,31823,9,''),(263190,31823,10,''),(263191,31823,11,''),(263192,31823,12,''),(263193,31823,13,''),(263194,31823,14,''),(263195,31823,15,''),(263196,31823,16,''),(263197,31823,17,''),(263198,31823,18,''),(263199,31823,19,''),(263200,31823,20,''),(263201,31825,1,''),(263202,31825,2,''),(263203,31825,3,''),(263204,31825,4,''),(263205,31825,5,''),(263206,31825,6,''),(263207,31825,7,''),(263208,31825,8,''),(263209,31825,9,''),(263210,31825,10,''),(263211,31825,11,''),(263212,31825,12,''),(263213,31825,13,''),(263214,31825,14,''),(263215,31825,15,''),(263216,31825,16,''),(263217,31825,17,''),(263218,31825,18,''),(263219,31825,19,''),(263220,31825,20,''),(263221,31934,1,''),(263222,31934,2,''),(263223,31934,3,''),(263224,31934,4,''),(263225,31934,5,''),(263226,31934,6,''),(263227,31934,7,''),(263228,31934,8,''),(263229,31934,9,''),(263230,31934,10,''),(263231,31934,11,''),(263232,31934,12,''),(263233,31934,13,''),(263234,31934,14,''),(263235,31934,15,''),(263236,31934,16,''),(263237,31934,17,''),(263238,31934,18,''),(263239,31934,19,''),(263240,31934,20,''),(263241,32018,1,''),(263242,32018,2,''),(263243,32018,3,''),(263244,32018,4,''),(263245,32018,5,''),(263246,32018,6,''),(263247,32018,7,''),(263248,32018,8,''),(263249,32018,9,''),(263250,32018,10,''),(263251,32018,11,''),(263252,32018,12,''),(263253,32018,13,''),(263254,32018,14,''),(263255,32018,15,''),(263256,32018,16,''),(263257,32018,17,''),(263258,32018,18,''),(263259,32018,19,''),(263260,32018,20,''),(263261,32050,1,''),(263262,32050,2,''),(263263,32050,3,''),(263264,32050,4,''),(263265,32050,5,''),(263266,32050,6,''),(263267,32050,7,''),(263268,32050,8,''),(263269,32050,9,''),(263270,32050,10,''),(263271,32050,11,''),(263272,32050,12,''),(263273,32050,13,''),(263274,32050,14,''),(263275,32050,15,''),(263276,32050,16,''),(263277,32050,17,''),(263278,32050,18,''),(263279,32050,19,''),(263280,32050,20,''),(263281,32111,1,''),(263282,32111,2,''),(263283,32111,3,''),(263284,32111,4,''),(263285,32111,5,''),(263286,32111,6,''),(263287,32111,7,''),(263288,32111,8,''),(263289,32111,9,''),(263290,32111,10,''),(263291,32111,11,''),(263292,32111,12,''),(263293,32111,13,''),(263294,32111,14,''),(263295,32111,15,''),(263296,32111,16,''),(263297,32111,17,''),(263298,32111,18,''),(263299,32111,19,''),(263300,32111,20,''),(263301,32112,1,''),(263302,32112,2,''),(263303,32112,3,''),(263304,32112,4,''),(263305,32112,5,''),(263306,32112,6,''),(263307,32112,7,''),(263308,32112,8,''),(263309,32112,9,''),(263310,32112,10,''),(263311,32112,11,''),(263312,32112,12,''),(263313,32112,13,''),(263314,32112,14,''),(263315,32112,15,''),(263316,32112,16,''),(263317,32112,17,''),(263318,32112,18,''),(263319,32112,19,''),(263320,32112,20,''),(263321,32327,1,''),(263322,32327,2,''),(263323,32327,3,''),(263324,32327,4,''),(263325,32327,5,''),(263326,32327,6,''),(263327,32327,7,''),(263328,32327,8,''),(263329,32327,9,''),(263330,32327,10,''),(263331,32327,11,''),(263332,32327,12,''),(263333,32327,13,''),(263334,32327,14,''),(263335,32327,15,''),(263336,32327,16,''),(263337,32327,17,''),(263338,32327,18,''),(263339,32327,19,''),(263340,32327,20,''),(263341,32328,1,''),(263342,32328,2,''),(263343,32328,3,''),(263344,32328,4,''),(263345,32328,5,''),(263346,32328,6,''),(263347,32328,7,''),(263348,32328,8,''),(263349,32328,9,''),(263350,32328,10,''),(263351,32328,11,''),(263352,32328,12,''),(263353,32328,13,''),(263354,32328,14,''),(263355,32328,15,''),(263356,32328,16,''),(263357,32328,17,''),(263358,32328,18,''),(263359,32328,19,''),(263360,32328,20,''),(263361,32420,1,''),(263362,32420,2,''),(263363,32420,3,''),(263364,32420,4,''),(263365,32420,5,''),(263366,32420,6,''),(263367,32420,7,''),(263368,32420,8,''),(263369,32420,9,''),(263370,32420,10,''),(263371,32420,11,''),(263372,32420,12,''),(263373,32420,13,''),(263374,32420,14,''),(263375,32420,15,''),(263376,32420,16,''),(263377,32420,17,''),(263378,32420,18,''),(263379,32420,19,''),(263380,32420,20,''),(263381,32457,1,''),(263382,32457,2,''),(263383,32457,3,''),(263384,32457,4,''),(263385,32457,5,''),(263386,32457,6,''),(263387,32457,7,''),(263388,32457,8,''),(263389,32457,9,''),(263390,32457,10,''),(263391,32457,11,''),(263392,32457,12,''),(263393,32457,13,''),(263394,32457,14,''),(263395,32457,15,''),(263396,32457,16,''),(263397,32457,17,''),(263398,32457,18,''),(263399,32457,19,''),(263400,32457,20,''),(263401,32458,1,''),(263402,32458,2,''),(263403,32458,3,''),(263404,32458,4,''),(263405,32458,5,''),(263406,32458,6,''),(263407,32458,7,''),(263408,32458,8,''),(263409,32458,9,''),(263410,32458,10,''),(263411,32458,11,''),(263412,32458,12,''),(263413,32458,13,''),(263414,32458,14,''),(263415,32458,15,''),(263416,32458,16,''),(263417,32458,17,''),(263418,32458,18,''),(263419,32458,19,''),(263420,32458,20,''),(263421,32499,1,''),(263422,32499,2,''),(263423,32499,3,''),(263424,32499,4,''),(263425,32499,5,''),(263426,32499,6,''),(263427,32499,7,''),(263428,32499,8,''),(263429,32499,9,''),(263430,32499,10,''),(263431,32499,11,''),(263432,32499,12,''),(263433,32499,13,''),(263434,32499,14,''),(263435,32499,15,''),(263436,32499,16,''),(263437,32499,17,''),(263438,32499,18,''),(263439,32499,19,''),(263440,32499,20,''),(263441,32696,1,''),(263442,32696,2,''),(263443,32696,3,''),(263444,32696,4,''),(263445,32696,5,''),(263446,32696,6,''),(263447,32696,7,''),(263448,32696,8,''),(263449,32696,9,''),(263450,32696,10,''),(263451,32696,11,''),(263452,32696,12,''),(263453,32696,13,''),(263454,32696,14,''),(263455,32696,15,''),(263456,32696,16,''),(263457,32696,17,''),(263458,32696,18,''),(263459,32696,19,''),(263460,32696,20,''),(263461,33193,1,''),(263462,33193,2,''),(263463,33193,3,''),(263464,33193,4,''),(263465,33193,5,''),(263466,33193,6,''),(263467,33193,7,''),(263468,33193,8,''),(263469,33193,9,''),(263470,33193,10,''),(263471,33193,11,''),(263472,33193,12,''),(263473,33193,13,''),(263474,33193,14,''),(263475,33193,15,''),(263476,33193,16,''),(263477,33193,17,''),(263478,33193,18,''),(263479,33193,19,''),(263480,33193,20,''),(263481,33252,1,''),(263482,33252,2,''),(263483,33252,3,''),(263484,33252,4,''),(263485,33252,5,''),(263486,33252,6,''),(263487,33252,7,''),(263488,33252,8,''),(263489,33252,9,''),(263490,33252,10,''),(263491,33252,11,''),(263492,33252,12,''),(263493,33252,13,''),(263494,33252,14,''),(263495,33252,15,''),(263496,33252,16,''),(263497,33252,17,''),(263498,33252,18,''),(263499,33252,19,''),(263500,33252,20,''),(263501,33255,1,''),(263502,33255,2,''),(263503,33255,3,''),(263504,33255,4,''),(263505,33255,5,''),(263506,33255,6,''),(263507,33255,7,''),(263508,33255,8,''),(263509,33255,9,''),(263510,33255,10,''),(263511,33255,11,''),(263512,33255,12,''),(263513,33255,13,''),(263514,33255,14,''),(263515,33255,15,''),(263516,33255,16,''),(263517,33255,17,''),(263518,33255,18,''),(263519,33255,19,''),(263520,33255,20,''),(263521,33309,1,''),(263522,33309,2,''),(263523,33309,3,''),(263524,33309,4,''),(263525,33309,5,''),(263526,33309,6,''),(263527,33309,7,''),(263528,33309,8,''),(263529,33309,9,''),(263530,33309,10,''),(263531,33309,11,''),(263532,33309,12,''),(263533,33309,13,''),(263534,33309,14,''),(263535,33309,15,''),(263536,33309,16,''),(263537,33309,17,''),(263538,33309,18,''),(263539,33309,19,''),(263540,33309,20,''),(263541,33330,1,''),(263542,33330,2,''),(263543,33330,3,''),(263544,33330,4,''),(263545,33330,5,''),(263546,33330,6,''),(263547,33330,7,''),(263548,33330,8,''),(263549,33330,9,''),(263550,33330,10,''),(263551,33330,11,''),(263552,33330,12,''),(263553,33330,13,''),(263554,33330,14,''),(263555,33330,15,''),(263556,33330,16,''),(263557,33330,17,''),(263558,33330,18,''),(263559,33330,19,''),(263560,33330,20,''),(263561,33416,1,''),(263562,33416,2,''),(263563,33416,3,''),(263564,33416,4,''),(263565,33416,5,''),(263566,33416,6,''),(263567,33416,7,''),(263568,33416,8,''),(263569,33416,9,''),(263570,33416,10,''),(263571,33416,11,''),(263572,33416,12,''),(263573,33416,13,''),(263574,33416,14,''),(263575,33416,15,''),(263576,33416,16,''),(263577,33416,17,''),(263578,33416,18,''),(263579,33416,19,''),(263580,33416,20,''),(263581,33615,1,''),(263582,33615,2,''),(263583,33615,3,''),(263584,33615,4,''),(263585,33615,5,''),(263586,33615,6,''),(263587,33615,7,''),(263588,33615,8,''),(263589,33615,9,''),(263590,33615,10,''),(263591,33615,11,''),(263592,33615,12,''),(263593,33615,13,''),(263594,33615,14,''),(263595,33615,15,''),(263596,33615,16,''),(263597,33615,17,''),(263598,33615,18,''),(263599,33615,19,''),(263600,33615,20,''),(263601,33690,1,''),(263602,33690,2,''),(263603,33690,3,''),(263604,33690,4,''),(263605,33690,5,''),(263606,33690,6,''),(263607,33690,7,''),(263608,33690,8,''),(263609,33690,9,''),(263610,33690,10,''),(263611,33690,11,''),(263612,33690,12,''),(263613,33690,13,''),(263614,33690,14,''),(263615,33690,15,''),(263616,33690,16,''),(263617,33690,17,''),(263618,33690,18,''),(263619,33690,19,''),(263620,33690,20,''),(263621,33937,1,''),(263622,33937,2,''),(263623,33937,3,''),(263624,33937,4,''),(263625,33937,5,''),(263626,33937,6,''),(263627,33937,7,''),(263628,33937,8,''),(263629,33937,9,''),(263630,33937,10,''),(263631,33937,11,''),(263632,33937,12,''),(263633,33937,13,''),(263634,33937,14,''),(263635,33937,15,''),(263636,33937,16,''),(263637,33937,17,''),(263638,33937,18,''),(263639,33937,19,''),(263640,33937,20,''),(263641,34316,1,''),(263642,34316,2,''),(263643,34316,3,''),(263644,34316,4,''),(263645,34316,5,''),(263646,34316,6,''),(263647,34316,7,''),(263648,34316,8,''),(263649,34316,9,''),(263650,34316,10,''),(263651,34316,11,''),(263652,34316,12,''),(263653,34316,13,''),(263654,34316,14,''),(263655,34316,15,''),(263656,34316,16,''),(263657,34316,17,''),(263658,34316,18,''),(263659,34316,19,''),(263660,34316,20,''),(263661,34331,1,''),(263662,34331,2,''),(263663,34331,3,''),(263664,34331,4,''),(263665,34331,5,''),(263666,34331,6,''),(263667,34331,7,''),(263668,34331,8,''),(263669,34331,9,''),(263670,34331,10,''),(263671,34331,11,''),(263672,34331,12,''),(263673,34331,13,''),(263674,34331,14,''),(263675,34331,15,''),(263676,34331,16,''),(263677,34331,17,''),(263678,34331,18,''),(263679,34331,19,''),(263680,34331,20,''),(263681,34446,1,''),(263682,34446,2,''),(263683,34446,3,''),(263684,34446,4,''),(263685,34446,5,''),(263686,34446,6,''),(263687,34446,7,''),(263688,34446,8,''),(263689,34446,9,''),(263690,34446,10,''),(263691,34446,11,''),(263692,34446,12,''),(263693,34446,13,''),(263694,34446,14,''),(263695,34446,15,''),(263696,34446,16,''),(263697,34446,17,''),(263698,34446,18,''),(263699,34446,19,''),(263700,34446,20,''),(263701,34456,1,''),(263702,34456,2,''),(263703,34456,3,''),(263704,34456,4,''),(263705,34456,5,''),(263706,34456,6,''),(263707,34456,7,''),(263708,34456,8,''),(263709,34456,9,''),(263710,34456,10,''),(263711,34456,11,''),(263712,34456,12,''),(263713,34456,13,''),(263714,34456,14,''),(263715,34456,15,''),(263716,34456,16,''),(263717,34456,17,''),(263718,34456,18,''),(263719,34456,19,''),(263720,34456,20,''),(263721,34517,1,''),(263722,34517,2,''),(263723,34517,3,''),(263724,34517,4,''),(263725,34517,5,''),(263726,34517,6,''),(263727,34517,7,''),(263728,34517,8,''),(263729,34517,9,''),(263730,34517,10,''),(263731,34517,11,''),(263732,34517,12,''),(263733,34517,13,''),(263734,34517,14,''),(263735,34517,15,''),(263736,34517,16,''),(263737,34517,17,''),(263738,34517,18,''),(263739,34517,19,''),(263740,34517,20,''),(263741,34548,1,''),(263742,34548,2,''),(263743,34548,3,''),(263744,34548,4,''),(263745,34548,5,''),(263746,34548,6,''),(263747,34548,7,''),(263748,34548,8,''),(263749,34548,9,''),(263750,34548,10,''),(263751,34548,11,''),(263752,34548,12,''),(263753,34548,13,''),(263754,34548,14,''),(263755,34548,15,''),(263756,34548,16,''),(263757,34548,17,''),(263758,34548,18,''),(263759,34548,19,''),(263760,34548,20,''),(263761,34583,1,''),(263762,34583,2,''),(263763,34583,3,''),(263764,34583,4,''),(263765,34583,5,''),(263766,34583,6,''),(263767,34583,7,''),(263768,34583,8,''),(263769,34583,9,''),(263770,34583,10,''),(263771,34583,11,''),(263772,34583,12,''),(263773,34583,13,''),(263774,34583,14,''),(263775,34583,15,''),(263776,34583,16,''),(263777,34583,17,''),(263778,34583,18,''),(263779,34583,19,''),(263780,34583,20,''),(263781,34587,1,''),(263782,34587,2,''),(263783,34587,3,''),(263784,34587,4,''),(263785,34587,5,''),(263786,34587,6,''),(263787,34587,7,''),(263788,34587,8,''),(263789,34587,9,''),(263790,34587,10,''),(263791,34587,11,''),(263792,34587,12,''),(263793,34587,13,''),(263794,34587,14,''),(263795,34587,15,''),(263796,34587,16,''),(263797,34587,17,''),(263798,34587,18,''),(263799,34587,19,''),(263800,34587,20,''),(263801,34611,1,''),(263802,34611,2,''),(263803,34611,3,''),(263804,34611,4,''),(263805,34611,5,''),(263806,34611,6,''),(263807,34611,7,''),(263808,34611,8,''),(263809,34611,9,''),(263810,34611,10,''),(263811,34611,11,''),(263812,34611,12,''),(263813,34611,13,''),(263814,34611,14,''),(263815,34611,15,''),(263816,34611,16,''),(263817,34611,17,''),(263818,34611,18,''),(263819,34611,19,''),(263820,34611,20,''),(263821,34639,1,''),(263822,34639,2,''),(263823,34639,3,''),(263824,34639,4,''),(263825,34639,5,''),(263826,34639,6,''),(263827,34639,7,''),(263828,34639,8,''),(263829,34639,9,''),(263830,34639,10,''),(263831,34639,11,''),(263832,34639,12,''),(263833,34639,13,''),(263834,34639,14,''),(263835,34639,15,''),(263836,34639,16,''),(263837,34639,17,''),(263838,34639,18,''),(263839,34639,19,''),(263840,34639,20,''),(263841,34642,1,''),(263842,34642,2,''),(263843,34642,3,''),(263844,34642,4,''),(263845,34642,5,''),(263846,34642,6,''),(263847,34642,7,''),(263848,34642,8,''),(263849,34642,9,''),(263850,34642,10,''),(263851,34642,11,''),(263852,34642,12,''),(263853,34642,13,''),(263854,34642,14,''),(263855,34642,15,''),(263856,34642,16,''),(263857,34642,17,''),(263858,34642,18,''),(263859,34642,19,''),(263860,34642,20,''),(263861,34670,1,''),(263862,34670,2,''),(263863,34670,3,''),(263864,34670,4,''),(263865,34670,5,''),(263866,34670,6,''),(263867,34670,7,''),(263868,34670,8,''),(263869,34670,9,''),(263870,34670,10,''),(263871,34670,11,''),(263872,34670,12,''),(263873,34670,13,''),(263874,34670,14,''),(263875,34670,15,''),(263876,34670,16,''),(263877,34670,17,''),(263878,34670,18,''),(263879,34670,19,''),(263880,34670,20,''),(263881,34672,1,''),(263882,34672,2,''),(263883,34672,3,''),(263884,34672,4,''),(263885,34672,5,''),(263886,34672,6,''),(263887,34672,7,''),(263888,34672,8,''),(263889,34672,9,''),(263890,34672,10,''),(263891,34672,11,''),(263892,34672,12,''),(263893,34672,13,''),(263894,34672,14,''),(263895,34672,15,''),(263896,34672,16,''),(263897,34672,17,''),(263898,34672,18,''),(263899,34672,19,''),(263900,34672,20,''),(263901,34681,1,''),(263902,34681,2,''),(263903,34681,3,''),(263904,34681,4,''),(263905,34681,5,''),(263906,34681,6,''),(263907,34681,7,''),(263908,34681,8,''),(263909,34681,9,''),(263910,34681,10,''),(263911,34681,11,''),(263912,34681,12,''),(263913,34681,13,''),(263914,34681,14,''),(263915,34681,15,''),(263916,34681,16,''),(263917,34681,17,''),(263918,34681,18,''),(263919,34681,19,''),(263920,34681,20,''),(263921,34682,1,''),(263922,34682,2,''),(263923,34682,3,''),(263924,34682,4,''),(263925,34682,5,''),(263926,34682,6,''),(263927,34682,7,''),(263928,34682,8,''),(263929,34682,9,''),(263930,34682,10,''),(263931,34682,11,''),(263932,34682,12,''),(263933,34682,13,''),(263934,34682,14,''),(263935,34682,15,''),(263936,34682,16,''),(263937,34682,17,''),(263938,34682,18,''),(263939,34682,19,''),(263940,34682,20,''),(263941,34836,1,''),(263942,34836,2,''),(263943,34836,3,''),(263944,34836,4,''),(263945,34836,5,''),(263946,34836,6,''),(263947,34836,7,''),(263948,34836,8,''),(263949,34836,9,''),(263950,34836,10,''),(263951,34836,11,''),(263952,34836,12,''),(263953,34836,13,''),(263954,34836,14,''),(263955,34836,15,''),(263956,34836,16,''),(263957,34836,17,''),(263958,34836,18,''),(263959,34836,19,''),(263960,34836,20,''),(263961,34918,1,''),(263962,34918,2,''),(263963,34918,3,''),(263964,34918,4,''),(263965,34918,5,''),(263966,34918,6,''),(263967,34918,7,''),(263968,34918,8,''),(263969,34918,9,''),(263970,34918,10,''),(263971,34918,11,''),(263972,34918,12,''),(263973,34918,13,''),(263974,34918,14,''),(263975,34918,15,''),(263976,34918,16,''),(263977,34918,17,''),(263978,34918,18,''),(263979,34918,19,''),(263980,34918,20,''),(263981,34919,1,''),(263982,34919,2,''),(263983,34919,3,''),(263984,34919,4,''),(263985,34919,5,''),(263986,34919,6,''),(263987,34919,7,''),(263988,34919,8,''),(263989,34919,9,''),(263990,34919,10,''),(263991,34919,11,''),(263992,34919,12,''),(263993,34919,13,''),(263994,34919,14,''),(263995,34919,15,''),(263996,34919,16,''),(263997,34919,17,''),(263998,34919,18,''),(263999,34919,19,''),(264000,34919,20,''),(264001,34975,1,''),(264002,34975,2,''),(264003,34975,3,''),(264004,34975,4,''),(264005,34975,5,''),(264006,34975,6,''),(264007,34975,7,''),(264008,34975,8,''),(264009,34975,9,''),(264010,34975,10,''),(264011,34975,11,''),(264012,34975,12,''),(264013,34975,13,''),(264014,34975,14,''),(264015,34975,15,''),(264016,34975,16,''),(264017,34975,17,''),(264018,34975,18,''),(264019,34975,19,''),(264020,34975,20,''),(264021,35111,1,''),(264022,35111,2,''),(264023,35111,3,''),(264024,35111,4,''),(264025,35111,5,''),(264026,35111,6,''),(264027,35111,7,''),(264028,35111,8,''),(264029,35111,9,''),(264030,35111,10,''),(264031,35111,11,''),(264032,35111,12,''),(264033,35111,13,''),(264034,35111,14,''),(264035,35111,15,''),(264036,35111,16,''),(264037,35111,17,''),(264038,35111,18,''),(264039,35111,19,''),(264040,35111,20,''),(264041,35112,1,''),(264042,35112,2,''),(264043,35112,3,''),(264044,35112,4,''),(264045,35112,5,''),(264046,35112,6,''),(264047,35112,7,''),(264048,35112,8,''),(264049,35112,9,''),(264050,35112,10,''),(264051,35112,11,''),(264052,35112,12,''),(264053,35112,13,''),(264054,35112,14,''),(264055,35112,15,''),(264056,35112,16,''),(264057,35112,17,''),(264058,35112,18,''),(264059,35112,19,''),(264060,35112,20,''),(264061,35199,1,''),(264062,35199,2,''),(264063,35199,3,''),(264064,35199,4,''),(264065,35199,5,''),(264066,35199,6,''),(264067,35199,7,''),(264068,35199,8,''),(264069,35199,9,''),(264070,35199,10,''),(264071,35199,11,''),(264072,35199,12,''),(264073,35199,13,''),(264074,35199,14,''),(264075,35199,15,''),(264076,35199,16,''),(264077,35199,17,''),(264078,35199,18,''),(264079,35199,19,''),(264080,35199,20,''),(264081,35260,1,''),(264082,35260,2,''),(264083,35260,3,''),(264084,35260,4,''),(264085,35260,5,''),(264086,35260,6,''),(264087,35260,7,''),(264088,35260,8,''),(264089,35260,9,''),(264090,35260,10,''),(264091,35260,11,''),(264092,35260,12,''),(264093,35260,13,''),(264094,35260,14,''),(264095,35260,15,''),(264096,35260,16,''),(264097,35260,17,''),(264098,35260,18,''),(264099,35260,19,''),(264100,35260,20,''),(264101,35273,1,''),(264102,35273,2,''),(264103,35273,3,''),(264104,35273,4,''),(264105,35273,5,''),(264106,35273,6,''),(264107,35273,7,''),(264108,35273,8,''),(264109,35273,9,''),(264110,35273,10,''),(264111,35273,11,''),(264112,35273,12,''),(264113,35273,13,''),(264114,35273,14,''),(264115,35273,15,''),(264116,35273,16,''),(264117,35273,17,''),(264118,35273,18,''),(264119,35273,19,''),(264120,35273,20,''),(264121,35274,1,''),(264122,35274,2,''),(264123,35274,3,''),(264124,35274,4,''),(264125,35274,5,''),(264126,35274,6,''),(264127,35274,7,''),(264128,35274,8,''),(264129,35274,9,''),(264130,35274,10,''),(264131,35274,11,''),(264132,35274,12,''),(264133,35274,13,''),(264134,35274,14,''),(264135,35274,15,''),(264136,35274,16,''),(264137,35274,17,''),(264138,35274,18,''),(264139,35274,19,''),(264140,35274,20,''),(264141,35328,1,''),(264142,35328,2,''),(264143,35328,3,''),(264144,35328,4,''),(264145,35328,5,''),(264146,35328,6,''),(264147,35328,7,''),(264148,35328,8,''),(264149,35328,9,''),(264150,35328,10,''),(264151,35328,11,''),(264152,35328,12,''),(264153,35328,13,''),(264154,35328,14,''),(264155,35328,15,''),(264156,35328,16,''),(264157,35328,17,''),(264158,35328,18,''),(264159,35328,19,''),(264160,35328,20,''),(264161,35340,1,''),(264162,35340,2,''),(264163,35340,3,''),(264164,35340,4,''),(264165,35340,5,''),(264166,35340,6,''),(264167,35340,7,''),(264168,35340,8,''),(264169,35340,9,''),(264170,35340,10,''),(264171,35340,11,''),(264172,35340,12,''),(264173,35340,13,''),(264174,35340,14,''),(264175,35340,15,''),(264176,35340,16,''),(264177,35340,17,''),(264178,35340,18,''),(264179,35340,19,''),(264180,35340,20,''),(264181,35378,1,''),(264182,35378,2,''),(264183,35378,3,''),(264184,35378,4,''),(264185,35378,5,''),(264186,35378,6,''),(264187,35378,7,''),(264188,35378,8,''),(264189,35378,9,''),(264190,35378,10,''),(264191,35378,11,''),(264192,35378,12,''),(264193,35378,13,''),(264194,35378,14,''),(264195,35378,15,''),(264196,35378,16,''),(264197,35378,17,''),(264198,35378,18,''),(264199,35378,19,''),(264200,35378,20,''),(264201,35380,1,''),(264202,35380,2,''),(264203,35380,3,''),(264204,35380,4,''),(264205,35380,5,''),(264206,35380,6,''),(264207,35380,7,''),(264208,35380,8,''),(264209,35380,9,''),(264210,35380,10,''),(264211,35380,11,''),(264212,35380,12,''),(264213,35380,13,''),(264214,35380,14,''),(264215,35380,15,''),(264216,35380,16,''),(264217,35380,17,''),(264218,35380,18,''),(264219,35380,19,''),(264220,35380,20,''),(264221,35394,1,''),(264222,35394,2,''),(264223,35394,3,''),(264224,35394,4,''),(264225,35394,5,''),(264226,35394,6,''),(264227,35394,7,''),(264228,35394,8,''),(264229,35394,9,''),(264230,35394,10,''),(264231,35394,11,''),(264232,35394,12,''),(264233,35394,13,''),(264234,35394,14,''),(264235,35394,15,''),(264236,35394,16,''),(264237,35394,17,''),(264238,35394,18,''),(264239,35394,19,''),(264240,35394,20,''),(264241,35396,1,''),(264242,35396,2,''),(264243,35396,3,''),(264244,35396,4,''),(264245,35396,5,''),(264246,35396,6,''),(264247,35396,7,''),(264248,35396,8,''),(264249,35396,9,''),(264250,35396,10,''),(264251,35396,11,''),(264252,35396,12,''),(264253,35396,13,''),(264254,35396,14,''),(264255,35396,15,''),(264256,35396,16,''),(264257,35396,17,''),(264258,35396,18,''),(264259,35396,19,''),(264260,35396,20,''),(264261,35588,1,''),(264262,35588,2,''),(264263,35588,3,''),(264264,35588,4,''),(264265,35588,5,''),(264266,35588,6,''),(264267,35588,7,''),(264268,35588,8,''),(264269,35588,9,''),(264270,35588,10,''),(264271,35588,11,''),(264272,35588,12,''),(264273,35588,13,''),(264274,35588,14,''),(264275,35588,15,''),(264276,35588,16,''),(264277,35588,17,''),(264278,35588,18,''),(264279,35588,19,''),(264280,35588,20,''),(264281,35603,1,''),(264282,35603,2,''),(264283,35603,3,''),(264284,35603,4,''),(264285,35603,5,''),(264286,35603,6,''),(264287,35603,7,''),(264288,35603,8,''),(264289,35603,9,''),(264290,35603,10,''),(264291,35603,11,''),(264292,35603,12,''),(264293,35603,13,''),(264294,35603,14,''),(264295,35603,15,''),(264296,35603,16,''),(264297,35603,17,''),(264298,35603,18,''),(264299,35603,19,''),(264300,35603,20,''),(264301,35623,1,''),(264302,35623,2,''),(264303,35623,3,''),(264304,35623,4,''),(264305,35623,5,''),(264306,35623,6,''),(264307,35623,7,''),(264308,35623,8,''),(264309,35623,9,''),(264310,35623,10,''),(264311,35623,11,''),(264312,35623,12,''),(264313,35623,13,''),(264314,35623,14,''),(264315,35623,15,''),(264316,35623,16,''),(264317,35623,17,''),(264318,35623,18,''),(264319,35623,19,''),(264320,35623,20,''),(264321,35624,1,''),(264322,35624,2,''),(264323,35624,3,''),(264324,35624,4,''),(264325,35624,5,''),(264326,35624,6,''),(264327,35624,7,''),(264328,35624,8,''),(264329,35624,9,''),(264330,35624,10,''),(264331,35624,11,''),(264332,35624,12,''),(264333,35624,13,''),(264334,35624,14,''),(264335,35624,15,''),(264336,35624,16,''),(264337,35624,17,''),(264338,35624,18,''),(264339,35624,19,''),(264340,35624,20,''),(264341,35625,1,''),(264342,35625,2,''),(264343,35625,3,''),(264344,35625,4,''),(264345,35625,5,''),(264346,35625,6,''),(264347,35625,7,''),(264348,35625,8,''),(264349,35625,9,''),(264350,35625,10,''),(264351,35625,11,''),(264352,35625,12,''),(264353,35625,13,''),(264354,35625,14,''),(264355,35625,15,''),(264356,35625,16,''),(264357,35625,17,''),(264358,35625,18,''),(264359,35625,19,''),(264360,35625,20,''),(264361,35693,1,''),(264362,35693,2,''),(264363,35693,3,''),(264364,35693,4,''),(264365,35693,5,''),(264366,35693,6,''),(264367,35693,7,''),(264368,35693,8,''),(264369,35693,9,''),(264370,35693,10,''),(264371,35693,11,''),(264372,35693,12,''),(264373,35693,13,''),(264374,35693,14,''),(264375,35693,15,''),(264376,35693,16,''),(264377,35693,17,''),(264378,35693,18,''),(264379,35693,19,''),(264380,35693,20,''),(264381,35741,1,''),(264382,35741,2,''),(264383,35741,3,''),(264384,35741,4,''),(264385,35741,5,''),(264386,35741,6,''),(264387,35741,7,''),(264388,35741,8,''),(264389,35741,9,''),(264390,35741,10,''),(264391,35741,11,''),(264392,35741,12,''),(264393,35741,13,''),(264394,35741,14,''),(264395,35741,15,''),(264396,35741,16,''),(264397,35741,17,''),(264398,35741,18,''),(264399,35741,19,''),(264400,35741,20,''),(264401,35876,1,''),(264402,35876,2,''),(264403,35876,3,''),(264404,35876,4,''),(264405,35876,5,''),(264406,35876,6,''),(264407,35876,7,''),(264408,35876,8,''),(264409,35876,9,''),(264410,35876,10,''),(264411,35876,11,''),(264412,35876,12,''),(264413,35876,13,''),(264414,35876,14,''),(264415,35876,15,''),(264416,35876,16,''),(264417,35876,17,''),(264418,35876,18,''),(264419,35876,19,''),(264420,35876,20,''),(264421,35938,1,''),(264422,35938,2,''),(264423,35938,3,''),(264424,35938,4,''),(264425,35938,5,''),(264426,35938,6,''),(264427,35938,7,''),(264428,35938,8,''),(264429,35938,9,''),(264430,35938,10,''),(264431,35938,11,''),(264432,35938,12,''),(264433,35938,13,''),(264434,35938,14,''),(264435,35938,15,''),(264436,35938,16,''),(264437,35938,17,''),(264438,35938,18,''),(264439,35938,19,''),(264440,35938,20,''),(264441,35957,1,''),(264442,35957,2,''),(264443,35957,3,''),(264444,35957,4,''),(264445,35957,5,''),(264446,35957,6,''),(264447,35957,7,''),(264448,35957,8,''),(264449,35957,9,''),(264450,35957,10,''),(264451,35957,11,''),(264452,35957,12,''),(264453,35957,13,''),(264454,35957,14,''),(264455,35957,15,''),(264456,35957,16,''),(264457,35957,17,''),(264458,35957,18,''),(264459,35957,19,''),(264460,35957,20,''),(264461,35958,1,''),(264462,35958,2,''),(264463,35958,3,''),(264464,35958,4,''),(264465,35958,5,''),(264466,35958,6,''),(264467,35958,7,''),(264468,35958,8,''),(264469,35958,9,''),(264470,35958,10,''),(264471,35958,11,''),(264472,35958,12,''),(264473,35958,13,''),(264474,35958,14,''),(264475,35958,15,''),(264476,35958,16,''),(264477,35958,17,''),(264478,35958,18,''),(264479,35958,19,''),(264480,35958,20,''),(264481,35968,1,''),(264482,35968,2,''),(264483,35968,3,''),(264484,35968,4,''),(264485,35968,5,''),(264486,35968,6,''),(264487,35968,7,''),(264488,35968,8,''),(264489,35968,9,''),(264490,35968,10,''),(264491,35968,11,''),(264492,35968,12,''),(264493,35968,13,''),(264494,35968,14,''),(264495,35968,15,''),(264496,35968,16,''),(264497,35968,17,''),(264498,35968,18,''),(264499,35968,19,''),(264500,35968,20,''),(264501,35977,1,''),(264502,35977,2,''),(264503,35977,3,''),(264504,35977,4,''),(264505,35977,5,''),(264506,35977,6,''),(264507,35977,7,''),(264508,35977,8,''),(264509,35977,9,''),(264510,35977,10,''),(264511,35977,11,''),(264512,35977,12,''),(264513,35977,13,''),(264514,35977,14,''),(264515,35977,15,''),(264516,35977,16,''),(264517,35977,17,''),(264518,35977,18,''),(264519,35977,19,''),(264520,35977,20,''),(264521,35981,1,''),(264522,35981,2,''),(264523,35981,3,''),(264524,35981,4,''),(264525,35981,5,''),(264526,35981,6,''),(264527,35981,7,''),(264528,35981,8,''),(264529,35981,9,''),(264530,35981,10,''),(264531,35981,11,''),(264532,35981,12,''),(264533,35981,13,''),(264534,35981,14,''),(264535,35981,15,''),(264536,35981,16,''),(264537,35981,17,''),(264538,35981,18,''),(264539,35981,19,''),(264540,35981,20,''),(264541,36001,1,''),(264542,36001,2,''),(264543,36001,3,''),(264544,36001,4,''),(264545,36001,5,''),(264546,36001,6,''),(264547,36001,7,''),(264548,36001,8,''),(264549,36001,9,''),(264550,36001,10,''),(264551,36001,11,''),(264552,36001,12,''),(264553,36001,13,''),(264554,36001,14,''),(264555,36001,15,''),(264556,36001,16,''),(264557,36001,17,''),(264558,36001,18,''),(264559,36001,19,''),(264560,36001,20,''),(264561,36016,1,''),(264562,36016,2,''),(264563,36016,3,''),(264564,36016,4,''),(264565,36016,5,''),(264566,36016,6,''),(264567,36016,7,''),(264568,36016,8,''),(264569,36016,9,''),(264570,36016,10,''),(264571,36016,11,''),(264572,36016,12,''),(264573,36016,13,''),(264574,36016,14,''),(264575,36016,15,''),(264576,36016,16,''),(264577,36016,17,''),(264578,36016,18,''),(264579,36016,19,''),(264580,36016,20,''),(264581,36066,1,''),(264582,36066,2,''),(264583,36066,3,''),(264584,36066,4,''),(264585,36066,5,''),(264586,36066,6,''),(264587,36066,7,''),(264588,36066,8,''),(264589,36066,9,''),(264590,36066,10,''),(264591,36066,11,''),(264592,36066,12,''),(264593,36066,13,''),(264594,36066,14,''),(264595,36066,15,''),(264596,36066,16,''),(264597,36066,17,''),(264598,36066,18,''),(264599,36066,19,''),(264600,36066,20,''),(264601,36075,1,''),(264602,36075,2,''),(264603,36075,3,''),(264604,36075,4,''),(264605,36075,5,''),(264606,36075,6,''),(264607,36075,7,''),(264608,36075,8,''),(264609,36075,9,''),(264610,36075,10,''),(264611,36075,11,''),(264612,36075,12,''),(264613,36075,13,''),(264614,36075,14,''),(264615,36075,15,''),(264616,36075,16,''),(264617,36075,17,''),(264618,36075,18,''),(264619,36075,19,''),(264620,36075,20,''),(264621,36225,1,''),(264622,36225,2,''),(264623,36225,3,''),(264624,36225,4,''),(264625,36225,5,''),(264626,36225,6,''),(264627,36225,7,''),(264628,36225,8,''),(264629,36225,9,''),(264630,36225,10,''),(264631,36225,11,''),(264632,36225,12,''),(264633,36225,13,''),(264634,36225,14,''),(264635,36225,15,''),(264636,36225,16,''),(264637,36225,17,''),(264638,36225,18,''),(264639,36225,19,''),(264640,36225,20,''),(264641,36243,1,''),(264642,36243,2,''),(264643,36243,3,''),(264644,36243,4,''),(264645,36243,5,''),(264646,36243,6,''),(264647,36243,7,''),(264648,36243,8,''),(264649,36243,9,''),(264650,36243,10,''),(264651,36243,11,''),(264652,36243,12,''),(264653,36243,13,''),(264654,36243,14,''),(264655,36243,15,''),(264656,36243,16,''),(264657,36243,17,''),(264658,36243,18,''),(264659,36243,19,''),(264660,36243,20,''),(264661,36283,1,''),(264662,36283,2,''),(264663,36283,3,''),(264664,36283,4,''),(264665,36283,5,''),(264666,36283,6,''),(264667,36283,7,''),(264668,36283,8,''),(264669,36283,9,''),(264670,36283,10,''),(264671,36283,11,''),(264672,36283,12,''),(264673,36283,13,''),(264674,36283,14,''),(264675,36283,15,''),(264676,36283,16,''),(264677,36283,17,''),(264678,36283,18,''),(264679,36283,19,''),(264680,36283,20,''),(264681,36323,1,''),(264682,36323,2,''),(264683,36323,3,''),(264684,36323,4,''),(264685,36323,5,''),(264686,36323,6,''),(264687,36323,7,''),(264688,36323,8,''),(264689,36323,9,''),(264690,36323,10,''),(264691,36323,11,''),(264692,36323,12,''),(264693,36323,13,''),(264694,36323,14,''),(264695,36323,15,''),(264696,36323,16,''),(264697,36323,17,''),(264698,36323,18,''),(264699,36323,19,''),(264700,36323,20,''),(264701,36355,1,''),(264702,36355,2,''),(264703,36355,3,''),(264704,36355,4,''),(264705,36355,5,''),(264706,36355,6,''),(264707,36355,7,''),(264708,36355,8,''),(264709,36355,9,''),(264710,36355,10,''),(264711,36355,11,''),(264712,36355,12,''),(264713,36355,13,''),(264714,36355,14,''),(264715,36355,15,''),(264716,36355,16,''),(264717,36355,17,''),(264718,36355,18,''),(264719,36355,19,''),(264720,36355,20,''),(264721,36360,1,''),(264722,36360,2,''),(264723,36360,3,''),(264724,36360,4,''),(264725,36360,5,''),(264726,36360,6,''),(264727,36360,7,''),(264728,36360,8,''),(264729,36360,9,''),(264730,36360,10,''),(264731,36360,11,''),(264732,36360,12,''),(264733,36360,13,''),(264734,36360,14,''),(264735,36360,15,''),(264736,36360,16,''),(264737,36360,17,''),(264738,36360,18,''),(264739,36360,19,''),(264740,36360,20,''),(264741,36405,1,''),(264742,36405,2,''),(264743,36405,3,''),(264744,36405,4,''),(264745,36405,5,''),(264746,36405,6,''),(264747,36405,7,''),(264748,36405,8,''),(264749,36405,9,''),(264750,36405,10,''),(264751,36405,11,''),(264752,36405,12,''),(264753,36405,13,''),(264754,36405,14,''),(264755,36405,15,''),(264756,36405,16,''),(264757,36405,17,''),(264758,36405,18,''),(264759,36405,19,''),(264760,36405,20,''),(264761,36411,1,''),(264762,36411,2,''),(264763,36411,3,''),(264764,36411,4,''),(264765,36411,5,''),(264766,36411,6,''),(264767,36411,7,''),(264768,36411,8,''),(264769,36411,9,''),(264770,36411,10,''),(264771,36411,11,''),(264772,36411,12,''),(264773,36411,13,''),(264774,36411,14,''),(264775,36411,15,''),(264776,36411,16,''),(264777,36411,17,''),(264778,36411,18,''),(264779,36411,19,''),(264780,36411,20,''),(264781,36412,1,''),(264782,36412,2,''),(264783,36412,3,''),(264784,36412,4,''),(264785,36412,5,''),(264786,36412,6,''),(264787,36412,7,''),(264788,36412,8,''),(264789,36412,9,''),(264790,36412,10,''),(264791,36412,11,''),(264792,36412,12,''),(264793,36412,13,''),(264794,36412,14,''),(264795,36412,15,''),(264796,36412,16,''),(264797,36412,17,''),(264798,36412,18,''),(264799,36412,19,''),(264800,36412,20,''),(264801,36423,1,''),(264802,36423,2,''),(264803,36423,3,''),(264804,36423,4,''),(264805,36423,5,''),(264806,36423,6,''),(264807,36423,7,''),(264808,36423,8,''),(264809,36423,9,''),(264810,36423,10,''),(264811,36423,11,''),(264812,36423,12,''),(264813,36423,13,''),(264814,36423,14,''),(264815,36423,15,''),(264816,36423,16,''),(264817,36423,17,''),(264818,36423,18,''),(264819,36423,19,''),(264820,36423,20,''),(264821,36454,1,''),(264822,36454,2,''),(264823,36454,3,''),(264824,36454,4,''),(264825,36454,5,''),(264826,36454,6,''),(264827,36454,7,''),(264828,36454,8,''),(264829,36454,9,''),(264830,36454,10,''),(264831,36454,11,''),(264832,36454,12,''),(264833,36454,13,''),(264834,36454,14,''),(264835,36454,15,''),(264836,36454,16,''),(264837,36454,17,''),(264838,36454,18,''),(264839,36454,19,''),(264840,36454,20,''),(264841,36519,1,''),(264842,36519,2,''),(264843,36519,3,''),(264844,36519,4,''),(264845,36519,5,''),(264846,36519,6,''),(264847,36519,7,''),(264848,36519,8,''),(264849,36519,9,''),(264850,36519,10,''),(264851,36519,11,''),(264852,36519,12,''),(264853,36519,13,''),(264854,36519,14,''),(264855,36519,15,''),(264856,36519,16,''),(264857,36519,17,''),(264858,36519,18,''),(264859,36519,19,''),(264860,36519,20,''),(264861,36575,1,''),(264862,36575,2,''),(264863,36575,3,''),(264864,36575,4,''),(264865,36575,5,''),(264866,36575,6,''),(264867,36575,7,''),(264868,36575,8,''),(264869,36575,9,''),(264870,36575,10,''),(264871,36575,11,''),(264872,36575,12,''),(264873,36575,13,''),(264874,36575,14,''),(264875,36575,15,''),(264876,36575,16,''),(264877,36575,17,''),(264878,36575,18,''),(264879,36575,19,''),(264880,36575,20,''),(264881,36576,1,''),(264882,36576,2,''),(264883,36576,3,''),(264884,36576,4,''),(264885,36576,5,''),(264886,36576,6,''),(264887,36576,7,''),(264888,36576,8,''),(264889,36576,9,''),(264890,36576,10,''),(264891,36576,11,''),(264892,36576,12,''),(264893,36576,13,''),(264894,36576,14,''),(264895,36576,15,''),(264896,36576,16,''),(264897,36576,17,''),(264898,36576,18,''),(264899,36576,19,''),(264900,36576,20,''),(264901,36577,1,''),(264902,36577,2,''),(264903,36577,3,''),(264904,36577,4,''),(264905,36577,5,''),(264906,36577,6,''),(264907,36577,7,''),(264908,36577,8,''),(264909,36577,9,''),(264910,36577,10,''),(264911,36577,11,''),(264912,36577,12,''),(264913,36577,13,''),(264914,36577,14,''),(264915,36577,15,''),(264916,36577,16,''),(264917,36577,17,''),(264918,36577,18,''),(264919,36577,19,''),(264920,36577,20,''),(264921,36578,1,''),(264922,36578,2,''),(264923,36578,3,''),(264924,36578,4,''),(264925,36578,5,''),(264926,36578,6,''),(264927,36578,7,''),(264928,36578,8,''),(264929,36578,9,''),(264930,36578,10,''),(264931,36578,11,''),(264932,36578,12,''),(264933,36578,13,''),(264934,36578,14,''),(264935,36578,15,''),(264936,36578,16,''),(264937,36578,17,''),(264938,36578,18,''),(264939,36578,19,''),(264940,36578,20,''),(264941,36636,1,''),(264942,36636,2,''),(264943,36636,3,''),(264944,36636,4,''),(264945,36636,5,''),(264946,36636,6,''),(264947,36636,7,''),(264948,36636,8,''),(264949,36636,9,''),(264950,36636,10,''),(264951,36636,11,''),(264952,36636,12,''),(264953,36636,13,''),(264954,36636,14,''),(264955,36636,15,''),(264956,36636,16,''),(264957,36636,17,''),(264958,36636,18,''),(264959,36636,19,''),(264960,36636,20,''),(264961,37280,1,''),(264962,37280,2,''),(264963,37280,3,''),(264964,37280,4,''),(264965,37280,5,''),(264966,37280,6,''),(264967,37280,7,''),(264968,37280,8,''),(264969,37280,9,''),(264970,37280,10,''),(264971,37280,11,''),(264972,37280,12,''),(264973,37280,13,''),(264974,37280,14,''),(264975,37280,15,''),(264976,37280,16,''),(264977,37280,17,''),(264978,37280,18,''),(264979,37280,19,''),(264980,37280,20,''),(264981,37287,1,''),(264982,37287,2,''),(264983,37287,3,''),(264984,37287,4,''),(264985,37287,5,''),(264986,37287,6,''),(264987,37287,7,''),(264988,37287,8,''),(264989,37287,9,''),(264990,37287,10,''),(264991,37287,11,''),(264992,37287,12,''),(264993,37287,13,''),(264994,37287,14,''),(264995,37287,15,''),(264996,37287,16,''),(264997,37287,17,''),(264998,37287,18,''),(264999,37287,19,''),(265000,37287,20,''),(265001,37314,1,''),(265002,37314,2,''),(265003,37314,3,''),(265004,37314,4,''),(265005,37314,5,''),(265006,37314,6,''),(265007,37314,7,''),(265008,37314,8,''),(265009,37314,9,''),(265010,37314,10,''),(265011,37314,11,''),(265012,37314,12,''),(265013,37314,13,''),(265014,37314,14,''),(265015,37314,15,''),(265016,37314,16,''),(265017,37314,17,''),(265018,37314,18,''),(265019,37314,19,''),(265020,37314,20,''),(265021,37357,1,''),(265022,37357,2,''),(265023,37357,3,''),(265024,37357,4,''),(265025,37357,5,''),(265026,37357,6,''),(265027,37357,7,''),(265028,37357,8,''),(265029,37357,9,''),(265030,37357,10,''),(265031,37357,11,''),(265032,37357,12,''),(265033,37357,13,''),(265034,37357,14,''),(265035,37357,15,''),(265036,37357,16,''),(265037,37357,17,''),(265038,37357,18,''),(265039,37357,19,''),(265040,37357,20,''),(265041,37368,1,''),(265042,37368,2,''),(265043,37368,3,''),(265044,37368,4,''),(265045,37368,5,''),(265046,37368,6,''),(265047,37368,7,''),(265048,37368,8,''),(265049,37368,9,''),(265050,37368,10,''),(265051,37368,11,''),(265052,37368,12,''),(265053,37368,13,''),(265054,37368,14,''),(265055,37368,15,''),(265056,37368,16,''),(265057,37368,17,''),(265058,37368,18,''),(265059,37368,19,''),(265060,37368,20,''),(265061,37372,1,''),(265062,37372,2,''),(265063,37372,3,''),(265064,37372,4,''),(265065,37372,5,''),(265066,37372,6,''),(265067,37372,7,''),(265068,37372,8,''),(265069,37372,9,''),(265070,37372,10,''),(265071,37372,11,''),(265072,37372,12,''),(265073,37372,13,''),(265074,37372,14,''),(265075,37372,15,''),(265076,37372,16,''),(265077,37372,17,''),(265078,37372,18,''),(265079,37372,19,''),(265080,37372,20,''),(265081,37460,1,''),(265082,37460,2,''),(265083,37460,3,''),(265084,37460,4,''),(265085,37460,5,''),(265086,37460,6,''),(265087,37460,7,''),(265088,37460,8,''),(265089,37460,9,''),(265090,37460,10,''),(265091,37460,11,''),(265092,37460,12,''),(265093,37460,13,''),(265094,37460,14,''),(265095,37460,15,''),(265096,37460,16,''),(265097,37460,17,''),(265098,37460,18,''),(265099,37460,19,''),(265100,37460,20,''),(265101,37499,1,''),(265102,37499,2,''),(265103,37499,3,''),(265104,37499,4,''),(265105,37499,5,''),(265106,37499,6,''),(265107,37499,7,''),(265108,37499,8,''),(265109,37499,9,''),(265110,37499,10,''),(265111,37499,11,''),(265112,37499,12,''),(265113,37499,13,''),(265114,37499,14,''),(265115,37499,15,''),(265116,37499,16,''),(265117,37499,17,''),(265118,37499,18,''),(265119,37499,19,''),(265120,37499,20,''),(265121,37501,1,''),(265122,37501,2,''),(265123,37501,3,''),(265124,37501,4,''),(265125,37501,5,''),(265126,37501,6,''),(265127,37501,7,''),(265128,37501,8,''),(265129,37501,9,''),(265130,37501,10,''),(265131,37501,11,''),(265132,37501,12,''),(265133,37501,13,''),(265134,37501,14,''),(265135,37501,15,''),(265136,37501,16,''),(265137,37501,17,''),(265138,37501,18,''),(265139,37501,19,''),(265140,37501,20,''),(265141,37568,1,''),(265142,37568,2,''),(265143,37568,3,''),(265144,37568,4,''),(265145,37568,5,''),(265146,37568,6,''),(265147,37568,7,''),(265148,37568,8,''),(265149,37568,9,''),(265150,37568,10,''),(265151,37568,11,''),(265152,37568,12,''),(265153,37568,13,''),(265154,37568,14,''),(265155,37568,15,''),(265156,37568,16,''),(265157,37568,17,''),(265158,37568,18,''),(265159,37568,19,''),(265160,37568,20,''),(265161,37571,1,''),(265162,37571,2,''),(265163,37571,3,''),(265164,37571,4,''),(265165,37571,5,''),(265166,37571,6,''),(265167,37571,7,''),(265168,37571,8,''),(265169,37571,9,''),(265170,37571,10,''),(265171,37571,11,''),(265172,37571,12,''),(265173,37571,13,''),(265174,37571,14,''),(265175,37571,15,''),(265176,37571,16,''),(265177,37571,17,''),(265178,37571,18,''),(265179,37571,19,''),(265180,37571,20,''),(265181,37588,1,''),(265182,37588,2,''),(265183,37588,3,''),(265184,37588,4,''),(265185,37588,5,''),(265186,37588,6,''),(265187,37588,7,''),(265188,37588,8,''),(265189,37588,9,''),(265190,37588,10,''),(265191,37588,11,''),(265192,37588,12,''),(265193,37588,13,''),(265194,37588,14,''),(265195,37588,15,''),(265196,37588,16,''),(265197,37588,17,''),(265198,37588,18,''),(265199,37588,19,''),(265200,37588,20,''),(265201,37637,1,''),(265202,37637,2,''),(265203,37637,3,''),(265204,37637,4,''),(265205,37637,5,''),(265206,37637,6,''),(265207,37637,7,''),(265208,37637,8,''),(265209,37637,9,''),(265210,37637,10,''),(265211,37637,11,''),(265212,37637,12,''),(265213,37637,13,''),(265214,37637,14,''),(265215,37637,15,''),(265216,37637,16,''),(265217,37637,17,''),(265218,37637,18,''),(265219,37637,19,''),(265220,37637,20,''),(265221,37682,1,''),(265222,37682,2,''),(265223,37682,3,''),(265224,37682,4,''),(265225,37682,5,''),(265226,37682,6,''),(265227,37682,7,''),(265228,37682,8,''),(265229,37682,9,''),(265230,37682,10,''),(265231,37682,11,''),(265232,37682,12,''),(265233,37682,13,''),(265234,37682,14,''),(265235,37682,15,''),(265236,37682,16,''),(265237,37682,17,''),(265238,37682,18,''),(265239,37682,19,''),(265240,37682,20,''),(265241,37732,1,''),(265242,37732,2,''),(265243,37732,3,''),(265244,37732,4,''),(265245,37732,5,''),(265246,37732,6,''),(265247,37732,7,''),(265248,37732,8,''),(265249,37732,9,''),(265250,37732,10,''),(265251,37732,11,''),(265252,37732,12,''),(265253,37732,13,''),(265254,37732,14,''),(265255,37732,15,''),(265256,37732,16,''),(265257,37732,17,''),(265258,37732,18,''),(265259,37732,19,''),(265260,37732,20,''),(265261,37757,1,''),(265262,37757,2,''),(265263,37757,3,''),(265264,37757,4,''),(265265,37757,5,''),(265266,37757,6,''),(265267,37757,7,''),(265268,37757,8,''),(265269,37757,9,''),(265270,37757,10,''),(265271,37757,11,''),(265272,37757,12,''),(265273,37757,13,''),(265274,37757,14,''),(265275,37757,15,''),(265276,37757,16,''),(265277,37757,17,''),(265278,37757,18,''),(265279,37757,19,''),(265280,37757,20,''),(265281,37775,1,''),(265282,37775,2,''),(265283,37775,3,''),(265284,37775,4,''),(265285,37775,5,''),(265286,37775,6,''),(265287,37775,7,''),(265288,37775,8,''),(265289,37775,9,''),(265290,37775,10,''),(265291,37775,11,''),(265292,37775,12,''),(265293,37775,13,''),(265294,37775,14,''),(265295,37775,15,''),(265296,37775,16,''),(265297,37775,17,''),(265298,37775,18,''),(265299,37775,19,''),(265300,37775,20,''),(265301,37797,1,''),(265302,37797,2,''),(265303,37797,3,''),(265304,37797,4,''),(265305,37797,5,''),(265306,37797,6,''),(265307,37797,7,''),(265308,37797,8,''),(265309,37797,9,''),(265310,37797,10,''),(265311,37797,11,''),(265312,37797,12,''),(265313,37797,13,''),(265314,37797,14,''),(265315,37797,15,''),(265316,37797,16,''),(265317,37797,17,''),(265318,37797,18,''),(265319,37797,19,''),(265320,37797,20,''),(265321,37798,1,''),(265322,37798,2,''),(265323,37798,3,''),(265324,37798,4,''),(265325,37798,5,''),(265326,37798,6,''),(265327,37798,7,''),(265328,37798,8,''),(265329,37798,9,''),(265330,37798,10,''),(265331,37798,11,''),(265332,37798,12,''),(265333,37798,13,''),(265334,37798,14,''),(265335,37798,15,''),(265336,37798,16,''),(265337,37798,17,''),(265338,37798,18,''),(265339,37798,19,''),(265340,37798,20,''),(265341,37800,1,''),(265342,37800,2,''),(265343,37800,3,''),(265344,37800,4,''),(265345,37800,5,''),(265346,37800,6,''),(265347,37800,7,''),(265348,37800,8,''),(265349,37800,9,''),(265350,37800,10,''),(265351,37800,11,''),(265352,37800,12,''),(265353,37800,13,''),(265354,37800,14,''),(265355,37800,15,''),(265356,37800,16,''),(265357,37800,17,''),(265358,37800,18,''),(265359,37800,19,''),(265360,37800,20,''),(265361,37852,1,''),(265362,37852,2,''),(265363,37852,3,''),(265364,37852,4,''),(265365,37852,5,''),(265366,37852,6,''),(265367,37852,7,''),(265368,37852,8,''),(265369,37852,9,''),(265370,37852,10,''),(265371,37852,11,''),(265372,37852,12,''),(265373,37852,13,''),(265374,37852,14,''),(265375,37852,15,''),(265376,37852,16,''),(265377,37852,17,''),(265378,37852,18,''),(265379,37852,19,''),(265380,37852,20,''),(265381,37874,1,''),(265382,37874,2,''),(265383,37874,3,''),(265384,37874,4,''),(265385,37874,5,''),(265386,37874,6,''),(265387,37874,7,''),(265388,37874,8,''),(265389,37874,9,''),(265390,37874,10,''),(265391,37874,11,''),(265392,37874,12,''),(265393,37874,13,''),(265394,37874,14,''),(265395,37874,15,''),(265396,37874,16,''),(265397,37874,17,''),(265398,37874,18,''),(265399,37874,19,''),(265400,37874,20,''),(265401,37950,1,''),(265402,37950,2,''),(265403,37950,3,''),(265404,37950,4,''),(265405,37950,5,''),(265406,37950,6,''),(265407,37950,7,''),(265408,37950,8,''),(265409,37950,9,''),(265410,37950,10,''),(265411,37950,11,''),(265412,37950,12,''),(265413,37950,13,''),(265414,37950,14,''),(265415,37950,15,''),(265416,37950,16,''),(265417,37950,17,''),(265418,37950,18,''),(265419,37950,19,''),(265420,37950,20,''),(265421,38000,1,''),(265422,38000,2,''),(265423,38000,3,''),(265424,38000,4,''),(265425,38000,5,''),(265426,38000,6,''),(265427,38000,7,''),(265428,38000,8,''),(265429,38000,9,''),(265430,38000,10,''),(265431,38000,11,''),(265432,38000,12,''),(265433,38000,13,''),(265434,38000,14,''),(265435,38000,15,''),(265436,38000,16,''),(265437,38000,17,''),(265438,38000,18,''),(265439,38000,19,''),(265440,38000,20,''),(265441,38004,1,''),(265442,38004,2,''),(265443,38004,3,''),(265444,38004,4,''),(265445,38004,5,''),(265446,38004,6,''),(265447,38004,7,''),(265448,38004,8,''),(265449,38004,9,''),(265450,38004,10,''),(265451,38004,11,''),(265452,38004,12,''),(265453,38004,13,''),(265454,38004,14,''),(265455,38004,15,''),(265456,38004,16,''),(265457,38004,17,''),(265458,38004,18,''),(265459,38004,19,''),(265460,38004,20,''),(265461,38005,1,''),(265462,38005,2,''),(265463,38005,3,''),(265464,38005,4,''),(265465,38005,5,''),(265466,38005,6,''),(265467,38005,7,''),(265468,38005,8,''),(265469,38005,9,''),(265470,38005,10,''),(265471,38005,11,''),(265472,38005,12,''),(265473,38005,13,''),(265474,38005,14,''),(265475,38005,15,''),(265476,38005,16,''),(265477,38005,17,''),(265478,38005,18,''),(265479,38005,19,''),(265480,38005,20,''),(265481,38034,1,''),(265482,38034,2,''),(265483,38034,3,''),(265484,38034,4,''),(265485,38034,5,''),(265486,38034,6,''),(265487,38034,7,''),(265488,38034,8,''),(265489,38034,9,''),(265490,38034,10,''),(265491,38034,11,''),(265492,38034,12,''),(265493,38034,13,''),(265494,38034,14,''),(265495,38034,15,''),(265496,38034,16,''),(265497,38034,17,''),(265498,38034,18,''),(265499,38034,19,''),(265500,38034,20,''),(265501,38035,1,''),(265502,38035,2,''),(265503,38035,3,''),(265504,38035,4,''),(265505,38035,5,''),(265506,38035,6,''),(265507,38035,7,''),(265508,38035,8,''),(265509,38035,9,''),(265510,38035,10,''),(265511,38035,11,''),(265512,38035,12,''),(265513,38035,13,''),(265514,38035,14,''),(265515,38035,15,''),(265516,38035,16,''),(265517,38035,17,''),(265518,38035,18,''),(265519,38035,19,''),(265520,38035,20,''),(265521,38045,1,''),(265522,38045,2,''),(265523,38045,3,''),(265524,38045,4,''),(265525,38045,5,''),(265526,38045,6,''),(265527,38045,7,''),(265528,38045,8,''),(265529,38045,9,''),(265530,38045,10,''),(265531,38045,11,''),(265532,38045,12,''),(265533,38045,13,''),(265534,38045,14,''),(265535,38045,15,''),(265536,38045,16,''),(265537,38045,17,''),(265538,38045,18,''),(265539,38045,19,''),(265540,38045,20,''),(265541,38046,1,''),(265542,38046,2,''),(265543,38046,3,''),(265544,38046,4,''),(265545,38046,5,''),(265546,38046,6,''),(265547,38046,7,''),(265548,38046,8,''),(265549,38046,9,''),(265550,38046,10,''),(265551,38046,11,''),(265552,38046,12,''),(265553,38046,13,''),(265554,38046,14,''),(265555,38046,15,''),(265556,38046,16,''),(265557,38046,17,''),(265558,38046,18,''),(265559,38046,19,''),(265560,38046,20,''),(265561,38047,1,''),(265562,38047,2,''),(265563,38047,3,''),(265564,38047,4,''),(265565,38047,5,''),(265566,38047,6,''),(265567,38047,7,''),(265568,38047,8,''),(265569,38047,9,''),(265570,38047,10,''),(265571,38047,11,''),(265572,38047,12,''),(265573,38047,13,''),(265574,38047,14,''),(265575,38047,15,''),(265576,38047,16,''),(265577,38047,17,''),(265578,38047,18,''),(265579,38047,19,''),(265580,38047,20,''),(265581,38048,1,''),(265582,38048,2,''),(265583,38048,3,''),(265584,38048,4,''),(265585,38048,5,''),(265586,38048,6,''),(265587,38048,7,''),(265588,38048,8,''),(265589,38048,9,''),(265590,38048,10,''),(265591,38048,11,''),(265592,38048,12,''),(265593,38048,13,''),(265594,38048,14,''),(265595,38048,15,''),(265596,38048,16,''),(265597,38048,17,''),(265598,38048,18,''),(265599,38048,19,''),(265600,38048,20,''),(265601,38121,1,''),(265602,38121,2,''),(265603,38121,3,''),(265604,38121,4,''),(265605,38121,5,''),(265606,38121,6,''),(265607,38121,7,''),(265608,38121,8,''),(265609,38121,9,''),(265610,38121,10,''),(265611,38121,11,''),(265612,38121,12,''),(265613,38121,13,''),(265614,38121,14,''),(265615,38121,15,''),(265616,38121,16,''),(265617,38121,17,''),(265618,38121,18,''),(265619,38121,19,''),(265620,38121,20,''),(265621,38122,1,''),(265622,38122,2,''),(265623,38122,3,''),(265624,38122,4,''),(265625,38122,5,''),(265626,38122,6,''),(265627,38122,7,''),(265628,38122,8,''),(265629,38122,9,''),(265630,38122,10,''),(265631,38122,11,''),(265632,38122,12,''),(265633,38122,13,''),(265634,38122,14,''),(265635,38122,15,''),(265636,38122,16,''),(265637,38122,17,''),(265638,38122,18,''),(265639,38122,19,''),(265640,38122,20,''),(265641,38123,1,''),(265642,38123,2,''),(265643,38123,3,''),(265644,38123,4,''),(265645,38123,5,''),(265646,38123,6,''),(265647,38123,7,''),(265648,38123,8,''),(265649,38123,9,''),(265650,38123,10,''),(265651,38123,11,''),(265652,38123,12,''),(265653,38123,13,''),(265654,38123,14,''),(265655,38123,15,''),(265656,38123,16,''),(265657,38123,17,''),(265658,38123,18,''),(265659,38123,19,''),(265660,38123,20,''),(265661,38124,1,''),(265662,38124,2,''),(265663,38124,3,''),(265664,38124,4,''),(265665,38124,5,''),(265666,38124,6,''),(265667,38124,7,''),(265668,38124,8,''),(265669,38124,9,''),(265670,38124,10,''),(265671,38124,11,''),(265672,38124,12,''),(265673,38124,13,''),(265674,38124,14,''),(265675,38124,15,''),(265676,38124,16,''),(265677,38124,17,''),(265678,38124,18,''),(265679,38124,19,''),(265680,38124,20,''),(265681,38276,1,''),(265682,38276,2,''),(265683,38276,3,''),(265684,38276,4,''),(265685,38276,5,''),(265686,38276,6,''),(265687,38276,7,''),(265688,38276,8,''),(265689,38276,9,''),(265690,38276,10,''),(265691,38276,11,''),(265692,38276,12,''),(265693,38276,13,''),(265694,38276,14,''),(265695,38276,15,''),(265696,38276,16,''),(265697,38276,17,''),(265698,38276,18,''),(265699,38276,19,''),(265700,38276,20,''),(265701,38330,1,''),(265702,38330,2,''),(265703,38330,3,''),(265704,38330,4,''),(265705,38330,5,''),(265706,38330,6,''),(265707,38330,7,''),(265708,38330,8,''),(265709,38330,9,''),(265710,38330,10,''),(265711,38330,11,''),(265712,38330,12,''),(265713,38330,13,''),(265714,38330,14,''),(265715,38330,15,''),(265716,38330,16,''),(265717,38330,17,''),(265718,38330,18,''),(265719,38330,19,''),(265720,38330,20,''),(265721,38420,1,''),(265722,38420,2,''),(265723,38420,3,''),(265724,38420,4,''),(265725,38420,5,''),(265726,38420,6,''),(265727,38420,7,''),(265728,38420,8,''),(265729,38420,9,''),(265730,38420,10,''),(265731,38420,11,''),(265732,38420,12,''),(265733,38420,13,''),(265734,38420,14,''),(265735,38420,15,''),(265736,38420,16,''),(265737,38420,17,''),(265738,38420,18,''),(265739,38420,19,''),(265740,38420,20,''),(265741,38477,1,''),(265742,38477,2,''),(265743,38477,3,''),(265744,38477,4,''),(265745,38477,5,''),(265746,38477,6,''),(265747,38477,7,''),(265748,38477,8,''),(265749,38477,9,''),(265750,38477,10,''),(265751,38477,11,''),(265752,38477,12,''),(265753,38477,13,''),(265754,38477,14,''),(265755,38477,15,''),(265756,38477,16,''),(265757,38477,17,''),(265758,38477,18,''),(265759,38477,19,''),(265760,38477,20,''),(265761,38501,1,''),(265762,38501,2,''),(265763,38501,3,''),(265764,38501,4,''),(265765,38501,5,''),(265766,38501,6,''),(265767,38501,7,''),(265768,38501,8,''),(265769,38501,9,''),(265770,38501,10,''),(265771,38501,11,''),(265772,38501,12,''),(265773,38501,13,''),(265774,38501,14,''),(265775,38501,15,''),(265776,38501,16,''),(265777,38501,17,''),(265778,38501,18,''),(265779,38501,19,''),(265780,38501,20,''),(265781,38547,1,''),(265782,38547,2,''),(265783,38547,3,''),(265784,38547,4,''),(265785,38547,5,''),(265786,38547,6,''),(265787,38547,7,''),(265788,38547,8,''),(265789,38547,9,''),(265790,38547,10,''),(265791,38547,11,''),(265792,38547,12,''),(265793,38547,13,''),(265794,38547,14,''),(265795,38547,15,''),(265796,38547,16,''),(265797,38547,17,''),(265798,38547,18,''),(265799,38547,19,''),(265800,38547,20,''),(265801,38548,1,''),(265802,38548,2,''),(265803,38548,3,''),(265804,38548,4,''),(265805,38548,5,''),(265806,38548,6,''),(265807,38548,7,''),(265808,38548,8,''),(265809,38548,9,''),(265810,38548,10,''),(265811,38548,11,''),(265812,38548,12,''),(265813,38548,13,''),(265814,38548,14,''),(265815,38548,15,''),(265816,38548,16,''),(265817,38548,17,''),(265818,38548,18,''),(265819,38548,19,''),(265820,38548,20,''),(265821,38571,1,''),(265822,38571,2,''),(265823,38571,3,''),(265824,38571,4,''),(265825,38571,5,''),(265826,38571,6,''),(265827,38571,7,''),(265828,38571,8,''),(265829,38571,9,''),(265830,38571,10,''),(265831,38571,11,''),(265832,38571,12,''),(265833,38571,13,''),(265834,38571,14,''),(265835,38571,15,''),(265836,38571,16,''),(265837,38571,17,''),(265838,38571,18,''),(265839,38571,19,''),(265840,38571,20,''),(265841,38700,1,''),(265842,38700,2,''),(265843,38700,3,''),(265844,38700,4,''),(265845,38700,5,''),(265846,38700,6,''),(265847,38700,7,''),(265848,38700,8,''),(265849,38700,9,''),(265850,38700,10,''),(265851,38700,11,''),(265852,38700,12,''),(265853,38700,13,''),(265854,38700,14,''),(265855,38700,15,''),(265856,38700,16,''),(265857,38700,17,''),(265858,38700,18,''),(265859,38700,19,''),(265860,38700,20,''),(265861,38701,1,''),(265862,38701,2,''),(265863,38701,3,''),(265864,38701,4,''),(265865,38701,5,''),(265866,38701,6,''),(265867,38701,7,''),(265868,38701,8,''),(265869,38701,9,''),(265870,38701,10,''),(265871,38701,11,''),(265872,38701,12,''),(265873,38701,13,''),(265874,38701,14,''),(265875,38701,15,''),(265876,38701,16,''),(265877,38701,17,''),(265878,38701,18,''),(265879,38701,19,''),(265880,38701,20,''),(265881,38710,1,''),(265882,38710,2,''),(265883,38710,3,''),(265884,38710,4,''),(265885,38710,5,''),(265886,38710,6,''),(265887,38710,7,''),(265888,38710,8,''),(265889,38710,9,''),(265890,38710,10,''),(265891,38710,11,''),(265892,38710,12,''),(265893,38710,13,''),(265894,38710,14,''),(265895,38710,15,''),(265896,38710,16,''),(265897,38710,17,''),(265898,38710,18,''),(265899,38710,19,''),(265900,38710,20,''),(265901,38744,1,''),(265902,38744,2,''),(265903,38744,3,''),(265904,38744,4,''),(265905,38744,5,''),(265906,38744,6,''),(265907,38744,7,''),(265908,38744,8,''),(265909,38744,9,''),(265910,38744,10,''),(265911,38744,11,''),(265912,38744,12,''),(265913,38744,13,''),(265914,38744,14,''),(265915,38744,15,''),(265916,38744,16,''),(265917,38744,17,''),(265918,38744,18,''),(265919,38744,19,''),(265920,38744,20,''),(265921,38767,1,''),(265922,38767,2,''),(265923,38767,3,''),(265924,38767,4,''),(265925,38767,5,''),(265926,38767,6,''),(265927,38767,7,''),(265928,38767,8,''),(265929,38767,9,''),(265930,38767,10,''),(265931,38767,11,''),(265932,38767,12,''),(265933,38767,13,''),(265934,38767,14,''),(265935,38767,15,''),(265936,38767,16,''),(265937,38767,17,''),(265938,38767,18,''),(265939,38767,19,''),(265940,38767,20,''),(265941,38768,1,''),(265942,38768,2,''),(265943,38768,3,''),(265944,38768,4,''),(265945,38768,5,''),(265946,38768,6,''),(265947,38768,7,''),(265948,38768,8,''),(265949,38768,9,''),(265950,38768,10,''),(265951,38768,11,''),(265952,38768,12,''),(265953,38768,13,''),(265954,38768,14,''),(265955,38768,15,''),(265956,38768,16,''),(265957,38768,17,''),(265958,38768,18,''),(265959,38768,19,''),(265960,38768,20,''),(265961,38769,1,''),(265962,38769,2,''),(265963,38769,3,''),(265964,38769,4,''),(265965,38769,5,''),(265966,38769,6,''),(265967,38769,7,''),(265968,38769,8,''),(265969,38769,9,''),(265970,38769,10,''),(265971,38769,11,''),(265972,38769,12,''),(265973,38769,13,''),(265974,38769,14,''),(265975,38769,15,''),(265976,38769,16,''),(265977,38769,17,''),(265978,38769,18,''),(265979,38769,19,''),(265980,38769,20,''),(265981,38770,1,''),(265982,38770,2,''),(265983,38770,3,''),(265984,38770,4,''),(265985,38770,5,''),(265986,38770,6,''),(265987,38770,7,''),(265988,38770,8,''),(265989,38770,9,''),(265990,38770,10,''),(265991,38770,11,''),(265992,38770,12,''),(265993,38770,13,''),(265994,38770,14,''),(265995,38770,15,''),(265996,38770,16,''),(265997,38770,17,''),(265998,38770,18,''),(265999,38770,19,''),(266000,38770,20,''),(266001,38776,1,''),(266002,38776,2,''),(266003,38776,3,''),(266004,38776,4,''),(266005,38776,5,''),(266006,38776,6,''),(266007,38776,7,''),(266008,38776,8,''),(266009,38776,9,''),(266010,38776,10,''),(266011,38776,11,''),(266012,38776,12,''),(266013,38776,13,''),(266014,38776,14,''),(266015,38776,15,''),(266016,38776,16,''),(266017,38776,17,''),(266018,38776,18,''),(266019,38776,19,''),(266020,38776,20,''),(266021,38785,1,''),(266022,38785,2,''),(266023,38785,3,''),(266024,38785,4,''),(266025,38785,5,''),(266026,38785,6,''),(266027,38785,7,''),(266028,38785,8,''),(266029,38785,9,''),(266030,38785,10,''),(266031,38785,11,''),(266032,38785,12,''),(266033,38785,13,''),(266034,38785,14,''),(266035,38785,15,''),(266036,38785,16,''),(266037,38785,17,''),(266038,38785,18,''),(266039,38785,19,''),(266040,38785,20,''),(266041,38841,1,''),(266042,38841,2,''),(266043,38841,3,''),(266044,38841,4,''),(266045,38841,5,''),(266046,38841,6,''),(266047,38841,7,''),(266048,38841,8,''),(266049,38841,9,''),(266050,38841,10,''),(266051,38841,11,''),(266052,38841,12,''),(266053,38841,13,''),(266054,38841,14,''),(266055,38841,15,''),(266056,38841,16,''),(266057,38841,17,''),(266058,38841,18,''),(266059,38841,19,''),(266060,38841,20,''),(266061,38971,1,''),(266062,38971,2,''),(266063,38971,3,''),(266064,38971,4,''),(266065,38971,5,''),(266066,38971,6,''),(266067,38971,7,''),(266068,38971,8,''),(266069,38971,9,''),(266070,38971,10,''),(266071,38971,11,''),(266072,38971,12,''),(266073,38971,13,''),(266074,38971,14,''),(266075,38971,15,''),(266076,38971,16,''),(266077,38971,17,''),(266078,38971,18,''),(266079,38971,19,''),(266080,38971,20,''),(266081,38972,1,''),(266082,38972,2,''),(266083,38972,3,''),(266084,38972,4,''),(266085,38972,5,''),(266086,38972,6,''),(266087,38972,7,''),(266088,38972,8,''),(266089,38972,9,''),(266090,38972,10,''),(266091,38972,11,''),(266092,38972,12,''),(266093,38972,13,''),(266094,38972,14,''),(266095,38972,15,''),(266096,38972,16,''),(266097,38972,17,''),(266098,38972,18,''),(266099,38972,19,''),(266100,38972,20,''),(266101,39061,1,''),(266102,39061,2,''),(266103,39061,3,''),(266104,39061,4,''),(266105,39061,5,''),(266106,39061,6,''),(266107,39061,7,''),(266108,39061,8,''),(266109,39061,9,''),(266110,39061,10,''),(266111,39061,11,''),(266112,39061,12,''),(266113,39061,13,''),(266114,39061,14,''),(266115,39061,15,''),(266116,39061,16,''),(266117,39061,17,''),(266118,39061,18,''),(266119,39061,19,''),(266120,39061,20,''),(266121,39062,1,''),(266122,39062,2,''),(266123,39062,3,''),(266124,39062,4,''),(266125,39062,5,''),(266126,39062,6,''),(266127,39062,7,''),(266128,39062,8,''),(266129,39062,9,''),(266130,39062,10,''),(266131,39062,11,''),(266132,39062,12,''),(266133,39062,13,''),(266134,39062,14,''),(266135,39062,15,''),(266136,39062,16,''),(266137,39062,17,''),(266138,39062,18,''),(266139,39062,19,''),(266140,39062,20,''),(266141,39071,1,''),(266142,39071,2,''),(266143,39071,3,''),(266144,39071,4,''),(266145,39071,5,''),(266146,39071,6,''),(266147,39071,7,''),(266148,39071,8,''),(266149,39071,9,''),(266150,39071,10,''),(266151,39071,11,''),(266152,39071,12,''),(266153,39071,13,''),(266154,39071,14,''),(266155,39071,15,''),(266156,39071,16,''),(266157,39071,17,''),(266158,39071,18,''),(266159,39071,19,''),(266160,39071,20,''),(266161,39072,1,''),(266162,39072,2,''),(266163,39072,3,''),(266164,39072,4,''),(266165,39072,5,''),(266166,39072,6,''),(266167,39072,7,''),(266168,39072,8,''),(266169,39072,9,''),(266170,39072,10,''),(266171,39072,11,''),(266172,39072,12,''),(266173,39072,13,''),(266174,39072,14,''),(266175,39072,15,''),(266176,39072,16,''),(266177,39072,17,''),(266178,39072,18,''),(266179,39072,19,''),(266180,39072,20,''),(266181,39102,1,''),(266182,39102,2,''),(266183,39102,3,''),(266184,39102,4,''),(266185,39102,5,''),(266186,39102,6,''),(266187,39102,7,''),(266188,39102,8,''),(266189,39102,9,''),(266190,39102,10,''),(266191,39102,11,''),(266192,39102,12,''),(266193,39102,13,''),(266194,39102,14,''),(266195,39102,15,''),(266196,39102,16,''),(266197,39102,17,''),(266198,39102,18,''),(266199,39102,19,''),(266200,39102,20,''),(266201,39121,1,''),(266202,39121,2,''),(266203,39121,3,''),(266204,39121,4,''),(266205,39121,5,''),(266206,39121,6,''),(266207,39121,7,''),(266208,39121,8,''),(266209,39121,9,''),(266210,39121,10,''),(266211,39121,11,''),(266212,39121,12,''),(266213,39121,13,''),(266214,39121,14,''),(266215,39121,15,''),(266216,39121,16,''),(266217,39121,17,''),(266218,39121,18,''),(266219,39121,19,''),(266220,39121,20,''),(266221,39123,1,''),(266222,39123,2,''),(266223,39123,3,''),(266224,39123,4,''),(266225,39123,5,''),(266226,39123,6,''),(266227,39123,7,''),(266228,39123,8,''),(266229,39123,9,''),(266230,39123,10,''),(266231,39123,11,''),(266232,39123,12,''),(266233,39123,13,''),(266234,39123,14,''),(266235,39123,15,''),(266236,39123,16,''),(266237,39123,17,''),(266238,39123,18,''),(266239,39123,19,''),(266240,39123,20,''),(266241,39136,1,''),(266242,39136,2,''),(266243,39136,3,''),(266244,39136,4,''),(266245,39136,5,''),(266246,39136,6,''),(266247,39136,7,''),(266248,39136,8,''),(266249,39136,9,''),(266250,39136,10,''),(266251,39136,11,''),(266252,39136,12,''),(266253,39136,13,''),(266254,39136,14,''),(266255,39136,15,''),(266256,39136,16,''),(266257,39136,17,''),(266258,39136,18,''),(266259,39136,19,''),(266260,39136,20,''),(266261,39137,1,''),(266262,39137,2,''),(266263,39137,3,''),(266264,39137,4,''),(266265,39137,5,''),(266266,39137,6,''),(266267,39137,7,''),(266268,39137,8,''),(266269,39137,9,''),(266270,39137,10,''),(266271,39137,11,''),(266272,39137,12,''),(266273,39137,13,''),(266274,39137,14,''),(266275,39137,15,''),(266276,39137,16,''),(266277,39137,17,''),(266278,39137,18,''),(266279,39137,19,''),(266280,39137,20,''),(266281,39138,1,''),(266282,39138,2,''),(266283,39138,3,''),(266284,39138,4,''),(266285,39138,5,''),(266286,39138,6,''),(266287,39138,7,''),(266288,39138,8,''),(266289,39138,9,''),(266290,39138,10,''),(266291,39138,11,''),(266292,39138,12,''),(266293,39138,13,''),(266294,39138,14,''),(266295,39138,15,''),(266296,39138,16,''),(266297,39138,17,''),(266298,39138,18,''),(266299,39138,19,''),(266300,39138,20,''),(266301,39199,1,''),(266302,39199,2,''),(266303,39199,3,''),(266304,39199,4,''),(266305,39199,5,''),(266306,39199,6,''),(266307,39199,7,''),(266308,39199,8,''),(266309,39199,9,''),(266310,39199,10,''),(266311,39199,11,''),(266312,39199,12,''),(266313,39199,13,''),(266314,39199,14,''),(266315,39199,15,''),(266316,39199,16,''),(266317,39199,17,''),(266318,39199,18,''),(266319,39199,19,''),(266320,39199,20,''),(266321,39217,1,''),(266322,39217,2,''),(266323,39217,3,''),(266324,39217,4,''),(266325,39217,5,''),(266326,39217,6,''),(266327,39217,7,''),(266328,39217,8,''),(266329,39217,9,''),(266330,39217,10,''),(266331,39217,11,''),(266332,39217,12,''),(266333,39217,13,''),(266334,39217,14,''),(266335,39217,15,''),(266336,39217,16,''),(266337,39217,17,''),(266338,39217,18,''),(266339,39217,19,''),(266340,39217,20,''),(266341,39218,1,''),(266342,39218,2,''),(266343,39218,3,''),(266344,39218,4,''),(266345,39218,5,''),(266346,39218,6,''),(266347,39218,7,''),(266348,39218,8,''),(266349,39218,9,''),(266350,39218,10,''),(266351,39218,11,''),(266352,39218,12,''),(266353,39218,13,''),(266354,39218,14,''),(266355,39218,15,''),(266356,39218,16,''),(266357,39218,17,''),(266358,39218,18,''),(266359,39218,19,''),(266360,39218,20,''),(266361,39220,1,''),(266362,39220,2,''),(266363,39220,3,''),(266364,39220,4,''),(266365,39220,5,''),(266366,39220,6,''),(266367,39220,7,''),(266368,39220,8,''),(266369,39220,9,''),(266370,39220,10,''),(266371,39220,11,''),(266372,39220,12,''),(266373,39220,13,''),(266374,39220,14,''),(266375,39220,15,''),(266376,39220,16,''),(266377,39220,17,''),(266378,39220,18,''),(266379,39220,19,''),(266380,39220,20,''),(266381,39408,1,''),(266382,39408,2,''),(266383,39408,3,''),(266384,39408,4,''),(266385,39408,5,''),(266386,39408,6,''),(266387,39408,7,''),(266388,39408,8,''),(266389,39408,9,''),(266390,39408,10,''),(266391,39408,11,''),(266392,39408,12,''),(266393,39408,13,''),(266394,39408,14,''),(266395,39408,15,''),(266396,39408,16,''),(266397,39408,17,''),(266398,39408,18,''),(266399,39408,19,''),(266400,39408,20,''),(266401,39410,1,''),(266402,39410,2,''),(266403,39410,3,''),(266404,39410,4,''),(266405,39410,5,''),(266406,39410,6,''),(266407,39410,7,''),(266408,39410,8,''),(266409,39410,9,''),(266410,39410,10,''),(266411,39410,11,''),(266412,39410,12,''),(266413,39410,13,''),(266414,39410,14,''),(266415,39410,15,''),(266416,39410,16,''),(266417,39410,17,''),(266418,39410,18,''),(266419,39410,19,''),(266420,39410,20,''),(266421,39421,1,''),(266422,39421,2,''),(266423,39421,3,''),(266424,39421,4,''),(266425,39421,5,''),(266426,39421,6,''),(266427,39421,7,''),(266428,39421,8,''),(266429,39421,9,''),(266430,39421,10,''),(266431,39421,11,''),(266432,39421,12,''),(266433,39421,13,''),(266434,39421,14,''),(266435,39421,15,''),(266436,39421,16,''),(266437,39421,17,''),(266438,39421,18,''),(266439,39421,19,''),(266440,39421,20,''),(266441,39447,1,''),(266442,39447,2,''),(266443,39447,3,''),(266444,39447,4,''),(266445,39447,5,''),(266446,39447,6,''),(266447,39447,7,''),(266448,39447,8,''),(266449,39447,9,''),(266450,39447,10,''),(266451,39447,11,''),(266452,39447,12,''),(266453,39447,13,''),(266454,39447,14,''),(266455,39447,15,''),(266456,39447,16,''),(266457,39447,17,''),(266458,39447,18,''),(266459,39447,19,''),(266460,39447,20,''),(266461,39450,1,''),(266462,39450,2,''),(266463,39450,3,''),(266464,39450,4,''),(266465,39450,5,''),(266466,39450,6,''),(266467,39450,7,''),(266468,39450,8,''),(266469,39450,9,''),(266470,39450,10,''),(266471,39450,11,''),(266472,39450,12,''),(266473,39450,13,''),(266474,39450,14,''),(266475,39450,15,''),(266476,39450,16,''),(266477,39450,17,''),(266478,39450,18,''),(266479,39450,19,''),(266480,39450,20,''),(266481,39451,1,''),(266482,39451,2,''),(266483,39451,3,''),(266484,39451,4,''),(266485,39451,5,''),(266486,39451,6,''),(266487,39451,7,''),(266488,39451,8,''),(266489,39451,9,''),(266490,39451,10,''),(266491,39451,11,''),(266492,39451,12,''),(266493,39451,13,''),(266494,39451,14,''),(266495,39451,15,''),(266496,39451,16,''),(266497,39451,17,''),(266498,39451,18,''),(266499,39451,19,''),(266500,39451,20,''),(266501,39456,1,''),(266502,39456,2,''),(266503,39456,3,''),(266504,39456,4,''),(266505,39456,5,''),(266506,39456,6,''),(266507,39456,7,''),(266508,39456,8,''),(266509,39456,9,''),(266510,39456,10,''),(266511,39456,11,''),(266512,39456,12,''),(266513,39456,13,''),(266514,39456,14,''),(266515,39456,15,''),(266516,39456,16,''),(266517,39456,17,''),(266518,39456,18,''),(266519,39456,19,''),(266520,39456,20,''),(266521,39483,1,''),(266522,39483,2,''),(266523,39483,3,''),(266524,39483,4,''),(266525,39483,5,''),(266526,39483,6,''),(266527,39483,7,''),(266528,39483,8,''),(266529,39483,9,''),(266530,39483,10,''),(266531,39483,11,''),(266532,39483,12,''),(266533,39483,13,''),(266534,39483,14,''),(266535,39483,15,''),(266536,39483,16,''),(266537,39483,17,''),(266538,39483,18,''),(266539,39483,19,''),(266540,39483,20,''),(266541,39484,1,''),(266542,39484,2,''),(266543,39484,3,''),(266544,39484,4,''),(266545,39484,5,''),(266546,39484,6,''),(266547,39484,7,''),(266548,39484,8,''),(266549,39484,9,''),(266550,39484,10,''),(266551,39484,11,''),(266552,39484,12,''),(266553,39484,13,''),(266554,39484,14,''),(266555,39484,15,''),(266556,39484,16,''),(266557,39484,17,''),(266558,39484,18,''),(266559,39484,19,''),(266560,39484,20,''),(266561,39505,1,''),(266562,39505,2,''),(266563,39505,3,''),(266564,39505,4,''),(266565,39505,5,''),(266566,39505,6,''),(266567,39505,7,''),(266568,39505,8,''),(266569,39505,9,''),(266570,39505,10,''),(266571,39505,11,''),(266572,39505,12,''),(266573,39505,13,''),(266574,39505,14,''),(266575,39505,15,''),(266576,39505,16,''),(266577,39505,17,''),(266578,39505,18,''),(266579,39505,19,''),(266580,39505,20,''),(266581,39568,1,''),(266582,39568,2,''),(266583,39568,3,''),(266584,39568,4,''),(266585,39568,5,''),(266586,39568,6,''),(266587,39568,7,''),(266588,39568,8,''),(266589,39568,9,''),(266590,39568,10,''),(266591,39568,11,''),(266592,39568,12,''),(266593,39568,13,''),(266594,39568,14,''),(266595,39568,15,''),(266596,39568,16,''),(266597,39568,17,''),(266598,39568,18,''),(266599,39568,19,''),(266600,39568,20,''),(266601,39569,1,''),(266602,39569,2,''),(266603,39569,3,''),(266604,39569,4,''),(266605,39569,5,''),(266606,39569,6,''),(266607,39569,7,''),(266608,39569,8,''),(266609,39569,9,''),(266610,39569,10,''),(266611,39569,11,''),(266612,39569,12,''),(266613,39569,13,''),(266614,39569,14,''),(266615,39569,15,''),(266616,39569,16,''),(266617,39569,17,''),(266618,39569,18,''),(266619,39569,19,''),(266620,39569,20,''),(266621,39570,1,''),(266622,39570,2,''),(266623,39570,3,''),(266624,39570,4,''),(266625,39570,5,''),(266626,39570,6,''),(266627,39570,7,''),(266628,39570,8,''),(266629,39570,9,''),(266630,39570,10,''),(266631,39570,11,''),(266632,39570,12,''),(266633,39570,13,''),(266634,39570,14,''),(266635,39570,15,''),(266636,39570,16,''),(266637,39570,17,''),(266638,39570,18,''),(266639,39570,19,''),(266640,39570,20,''),(266641,39571,1,''),(266642,39571,2,''),(266643,39571,3,''),(266644,39571,4,''),(266645,39571,5,''),(266646,39571,6,''),(266647,39571,7,''),(266648,39571,8,''),(266649,39571,9,''),(266650,39571,10,''),(266651,39571,11,''),(266652,39571,12,''),(266653,39571,13,''),(266654,39571,14,''),(266655,39571,15,''),(266656,39571,16,''),(266657,39571,17,''),(266658,39571,18,''),(266659,39571,19,''),(266660,39571,20,''),(266661,39572,1,''),(266662,39572,2,''),(266663,39572,3,''),(266664,39572,4,''),(266665,39572,5,''),(266666,39572,6,''),(266667,39572,7,''),(266668,39572,8,''),(266669,39572,9,''),(266670,39572,10,''),(266671,39572,11,''),(266672,39572,12,''),(266673,39572,13,''),(266674,39572,14,''),(266675,39572,15,''),(266676,39572,16,''),(266677,39572,17,''),(266678,39572,18,''),(266679,39572,19,''),(266680,39572,20,''),(266681,39638,1,''),(266682,39638,2,''),(266683,39638,3,''),(266684,39638,4,''),(266685,39638,5,''),(266686,39638,6,''),(266687,39638,7,''),(266688,39638,8,''),(266689,39638,9,''),(266690,39638,10,''),(266691,39638,11,''),(266692,39638,12,''),(266693,39638,13,''),(266694,39638,14,''),(266695,39638,15,''),(266696,39638,16,''),(266697,39638,17,''),(266698,39638,18,''),(266699,39638,19,''),(266700,39638,20,''),(266701,39639,1,''),(266702,39639,2,''),(266703,39639,3,''),(266704,39639,4,''),(266705,39639,5,''),(266706,39639,6,''),(266707,39639,7,''),(266708,39639,8,''),(266709,39639,9,''),(266710,39639,10,''),(266711,39639,11,''),(266712,39639,12,''),(266713,39639,13,''),(266714,39639,14,''),(266715,39639,15,''),(266716,39639,16,''),(266717,39639,17,''),(266718,39639,18,''),(266719,39639,19,''),(266720,39639,20,''),(266721,39640,1,''),(266722,39640,2,''),(266723,39640,3,''),(266724,39640,4,''),(266725,39640,5,''),(266726,39640,6,''),(266727,39640,7,''),(266728,39640,8,''),(266729,39640,9,''),(266730,39640,10,''),(266731,39640,11,''),(266732,39640,12,''),(266733,39640,13,''),(266734,39640,14,''),(266735,39640,15,''),(266736,39640,16,''),(266737,39640,17,''),(266738,39640,18,''),(266739,39640,19,''),(266740,39640,20,''),(266741,39681,1,''),(266742,39681,2,''),(266743,39681,3,''),(266744,39681,4,''),(266745,39681,5,''),(266746,39681,6,''),(266747,39681,7,''),(266748,39681,8,''),(266749,39681,9,''),(266750,39681,10,''),(266751,39681,11,''),(266752,39681,12,''),(266753,39681,13,''),(266754,39681,14,''),(266755,39681,15,''),(266756,39681,16,''),(266757,39681,17,''),(266758,39681,18,''),(266759,39681,19,''),(266760,39681,20,''),(266761,39682,1,''),(266762,39682,2,''),(266763,39682,3,''),(266764,39682,4,''),(266765,39682,5,''),(266766,39682,6,''),(266767,39682,7,''),(266768,39682,8,''),(266769,39682,9,''),(266770,39682,10,''),(266771,39682,11,''),(266772,39682,12,''),(266773,39682,13,''),(266774,39682,14,''),(266775,39682,15,''),(266776,39682,16,''),(266777,39682,17,''),(266778,39682,18,''),(266779,39682,19,''),(266780,39682,20,''),(266781,39688,1,''),(266782,39688,2,''),(266783,39688,3,''),(266784,39688,4,''),(266785,39688,5,''),(266786,39688,6,''),(266787,39688,7,''),(266788,39688,8,''),(266789,39688,9,''),(266790,39688,10,''),(266791,39688,11,''),(266792,39688,12,''),(266793,39688,13,''),(266794,39688,14,''),(266795,39688,15,''),(266796,39688,16,''),(266797,39688,17,''),(266798,39688,18,''),(266799,39688,19,''),(266800,39688,20,''),(266801,39689,1,''),(266802,39689,2,''),(266803,39689,3,''),(266804,39689,4,''),(266805,39689,5,''),(266806,39689,6,''),(266807,39689,7,''),(266808,39689,8,''),(266809,39689,9,''),(266810,39689,10,''),(266811,39689,11,''),(266812,39689,12,''),(266813,39689,13,''),(266814,39689,14,''),(266815,39689,15,''),(266816,39689,16,''),(266817,39689,17,''),(266818,39689,18,''),(266819,39689,19,''),(266820,39689,20,''),(266821,39690,1,''),(266822,39690,2,''),(266823,39690,3,''),(266824,39690,4,''),(266825,39690,5,''),(266826,39690,6,''),(266827,39690,7,''),(266828,39690,8,''),(266829,39690,9,''),(266830,39690,10,''),(266831,39690,11,''),(266832,39690,12,''),(266833,39690,13,''),(266834,39690,14,''),(266835,39690,15,''),(266836,39690,16,''),(266837,39690,17,''),(266838,39690,18,''),(266839,39690,19,''),(266840,39690,20,''),(266841,39738,1,''),(266842,39738,2,''),(266843,39738,3,''),(266844,39738,4,''),(266845,39738,5,''),(266846,39738,6,''),(266847,39738,7,''),(266848,39738,8,''),(266849,39738,9,''),(266850,39738,10,''),(266851,39738,11,''),(266852,39738,12,''),(266853,39738,13,''),(266854,39738,14,''),(266855,39738,15,''),(266856,39738,16,''),(266857,39738,17,''),(266858,39738,18,''),(266859,39738,19,''),(266860,39738,20,''),(266861,39771,1,''),(266862,39771,2,''),(266863,39771,3,''),(266864,39771,4,''),(266865,39771,5,''),(266866,39771,6,''),(266867,39771,7,''),(266868,39771,8,''),(266869,39771,9,''),(266870,39771,10,''),(266871,39771,11,''),(266872,39771,12,''),(266873,39771,13,''),(266874,39771,14,''),(266875,39771,15,''),(266876,39771,16,''),(266877,39771,17,''),(266878,39771,18,''),(266879,39771,19,''),(266880,39771,20,''),(266881,39849,1,''),(266882,39849,2,''),(266883,39849,3,''),(266884,39849,4,''),(266885,39849,5,''),(266886,39849,6,''),(266887,39849,7,''),(266888,39849,8,''),(266889,39849,9,''),(266890,39849,10,''),(266891,39849,11,''),(266892,39849,12,''),(266893,39849,13,''),(266894,39849,14,''),(266895,39849,15,''),(266896,39849,16,''),(266897,39849,17,''),(266898,39849,18,''),(266899,39849,19,''),(266900,39849,20,''),(266901,39854,1,''),(266902,39854,2,''),(266903,39854,3,''),(266904,39854,4,''),(266905,39854,5,''),(266906,39854,6,''),(266907,39854,7,''),(266908,39854,8,''),(266909,39854,9,''),(266910,39854,10,''),(266911,39854,11,''),(266912,39854,12,''),(266913,39854,13,''),(266914,39854,14,''),(266915,39854,15,''),(266916,39854,16,''),(266917,39854,17,''),(266918,39854,18,''),(266919,39854,19,''),(266920,39854,20,''),(266921,39855,1,''),(266922,39855,2,''),(266923,39855,3,''),(266924,39855,4,''),(266925,39855,5,''),(266926,39855,6,''),(266927,39855,7,''),(266928,39855,8,''),(266929,39855,9,''),(266930,39855,10,''),(266931,39855,11,''),(266932,39855,12,''),(266933,39855,13,''),(266934,39855,14,''),(266935,39855,15,''),(266936,39855,16,''),(266937,39855,17,''),(266938,39855,18,''),(266939,39855,19,''),(266940,39855,20,''),(266941,39856,1,''),(266942,39856,2,''),(266943,39856,3,''),(266944,39856,4,''),(266945,39856,5,''),(266946,39856,6,''),(266947,39856,7,''),(266948,39856,8,''),(266949,39856,9,''),(266950,39856,10,''),(266951,39856,11,''),(266952,39856,12,''),(266953,39856,13,''),(266954,39856,14,''),(266955,39856,15,''),(266956,39856,16,''),(266957,39856,17,''),(266958,39856,18,''),(266959,39856,19,''),(266960,39856,20,''),(266961,39976,1,''),(266962,39976,2,''),(266963,39976,3,''),(266964,39976,4,''),(266965,39976,5,''),(266966,39976,6,''),(266967,39976,7,''),(266968,39976,8,''),(266969,39976,9,''),(266970,39976,10,''),(266971,39976,11,''),(266972,39976,12,''),(266973,39976,13,''),(266974,39976,14,''),(266975,39976,15,''),(266976,39976,16,''),(266977,39976,17,''),(266978,39976,18,''),(266979,39976,19,''),(266980,39976,20,''),(266981,40023,1,''),(266982,40023,2,''),(266983,40023,3,''),(266984,40023,4,''),(266985,40023,5,''),(266986,40023,6,''),(266987,40023,7,''),(266988,40023,8,''),(266989,40023,9,''),(266990,40023,10,''),(266991,40023,11,''),(266992,40023,12,''),(266993,40023,13,''),(266994,40023,14,''),(266995,40023,15,''),(266996,40023,16,''),(266997,40023,17,''),(266998,40023,18,''),(266999,40023,19,''),(267000,40023,20,''),(267001,40074,1,''),(267002,40074,2,''),(267003,40074,3,''),(267004,40074,4,''),(267005,40074,5,''),(267006,40074,6,''),(267007,40074,7,''),(267008,40074,8,''),(267009,40074,9,''),(267010,40074,10,''),(267011,40074,11,''),(267012,40074,12,''),(267013,40074,13,''),(267014,40074,14,''),(267015,40074,15,''),(267016,40074,16,''),(267017,40074,17,''),(267018,40074,18,''),(267019,40074,19,''),(267020,40074,20,''),(267021,40150,1,''),(267022,40150,2,''),(267023,40150,3,''),(267024,40150,4,''),(267025,40150,5,''),(267026,40150,6,''),(267027,40150,7,''),(267028,40150,8,''),(267029,40150,9,''),(267030,40150,10,''),(267031,40150,11,''),(267032,40150,12,''),(267033,40150,13,''),(267034,40150,14,''),(267035,40150,15,''),(267036,40150,16,''),(267037,40150,17,''),(267038,40150,18,''),(267039,40150,19,''),(267040,40150,20,''),(267041,40151,1,''),(267042,40151,2,''),(267043,40151,3,''),(267044,40151,4,''),(267045,40151,5,''),(267046,40151,6,''),(267047,40151,7,''),(267048,40151,8,''),(267049,40151,9,''),(267050,40151,10,''),(267051,40151,11,''),(267052,40151,12,''),(267053,40151,13,''),(267054,40151,14,''),(267055,40151,15,''),(267056,40151,16,''),(267057,40151,17,''),(267058,40151,18,''),(267059,40151,19,''),(267060,40151,20,''),(267061,40152,1,''),(267062,40152,2,''),(267063,40152,3,''),(267064,40152,4,''),(267065,40152,5,''),(267066,40152,6,''),(267067,40152,7,''),(267068,40152,8,''),(267069,40152,9,''),(267070,40152,10,''),(267071,40152,11,''),(267072,40152,12,''),(267073,40152,13,''),(267074,40152,14,''),(267075,40152,15,''),(267076,40152,16,''),(267077,40152,17,''),(267078,40152,18,''),(267079,40152,19,''),(267080,40152,20,''),(267081,40153,1,''),(267082,40153,2,''),(267083,40153,3,''),(267084,40153,4,''),(267085,40153,5,''),(267086,40153,6,''),(267087,40153,7,''),(267088,40153,8,''),(267089,40153,9,''),(267090,40153,10,''),(267091,40153,11,''),(267092,40153,12,''),(267093,40153,13,''),(267094,40153,14,''),(267095,40153,15,''),(267096,40153,16,''),(267097,40153,17,''),(267098,40153,18,''),(267099,40153,19,''),(267100,40153,20,''),(267101,40268,1,''),(267102,40268,2,''),(267103,40268,3,''),(267104,40268,4,''),(267105,40268,5,''),(267106,40268,6,''),(267107,40268,7,''),(267108,40268,8,''),(267109,40268,9,''),(267110,40268,10,''),(267111,40268,11,''),(267112,40268,12,''),(267113,40268,13,''),(267114,40268,14,''),(267115,40268,15,''),(267116,40268,16,''),(267117,40268,17,''),(267118,40268,18,''),(267119,40268,19,''),(267120,40268,20,''),(267121,40320,1,''),(267122,40320,2,''),(267123,40320,3,''),(267124,40320,4,''),(267125,40320,5,''),(267126,40320,6,''),(267127,40320,7,''),(267128,40320,8,''),(267129,40320,9,''),(267130,40320,10,''),(267131,40320,11,''),(267132,40320,12,''),(267133,40320,13,''),(267134,40320,14,''),(267135,40320,15,''),(267136,40320,16,''),(267137,40320,17,''),(267138,40320,18,''),(267139,40320,19,''),(267140,40320,20,''),(267141,40322,1,''),(267142,40322,2,''),(267143,40322,3,''),(267144,40322,4,''),(267145,40322,5,''),(267146,40322,6,''),(267147,40322,7,''),(267148,40322,8,''),(267149,40322,9,''),(267150,40322,10,''),(267151,40322,11,''),(267152,40322,12,''),(267153,40322,13,''),(267154,40322,14,''),(267155,40322,15,''),(267156,40322,16,''),(267157,40322,17,''),(267158,40322,18,''),(267159,40322,19,''),(267160,40322,20,''),(267161,40350,1,''),(267162,40350,2,''),(267163,40350,3,''),(267164,40350,4,''),(267165,40350,5,''),(267166,40350,6,''),(267167,40350,7,''),(267168,40350,8,''),(267169,40350,9,''),(267170,40350,10,''),(267171,40350,11,''),(267172,40350,12,''),(267173,40350,13,''),(267174,40350,14,''),(267175,40350,15,''),(267176,40350,16,''),(267177,40350,17,''),(267178,40350,18,''),(267179,40350,19,''),(267180,40350,20,''),(267181,40385,1,''),(267182,40385,2,''),(267183,40385,3,''),(267184,40385,4,''),(267185,40385,5,''),(267186,40385,6,''),(267187,40385,7,''),(267188,40385,8,''),(267189,40385,9,''),(267190,40385,10,''),(267191,40385,11,''),(267192,40385,12,''),(267193,40385,13,''),(267194,40385,14,''),(267195,40385,15,''),(267196,40385,16,''),(267197,40385,17,''),(267198,40385,18,''),(267199,40385,19,''),(267200,40385,20,''),(267201,40447,1,''),(267202,40447,2,''),(267203,40447,3,''),(267204,40447,4,''),(267205,40447,5,''),(267206,40447,6,''),(267207,40447,7,''),(267208,40447,8,''),(267209,40447,9,''),(267210,40447,10,''),(267211,40447,11,''),(267212,40447,12,''),(267213,40447,13,''),(267214,40447,14,''),(267215,40447,15,''),(267216,40447,16,''),(267217,40447,17,''),(267218,40447,18,''),(267219,40447,19,''),(267220,40447,20,''),(267221,40453,1,''),(267222,40453,2,''),(267223,40453,3,''),(267224,40453,4,''),(267225,40453,5,''),(267226,40453,6,''),(267227,40453,7,''),(267228,40453,8,''),(267229,40453,9,''),(267230,40453,10,''),(267231,40453,11,''),(267232,40453,12,''),(267233,40453,13,''),(267234,40453,14,''),(267235,40453,15,''),(267236,40453,16,''),(267237,40453,17,''),(267238,40453,18,''),(267239,40453,19,''),(267240,40453,20,''),(267241,40459,1,''),(267242,40459,2,''),(267243,40459,3,''),(267244,40459,4,''),(267245,40459,5,''),(267246,40459,6,''),(267247,40459,7,''),(267248,40459,8,''),(267249,40459,9,''),(267250,40459,10,''),(267251,40459,11,''),(267252,40459,12,''),(267253,40459,13,''),(267254,40459,14,''),(267255,40459,15,''),(267256,40459,16,''),(267257,40459,17,''),(267258,40459,18,''),(267259,40459,19,''),(267260,40459,20,''),(267261,40495,1,''),(267262,40495,2,''),(267263,40495,3,''),(267264,40495,4,''),(267265,40495,5,''),(267266,40495,6,''),(267267,40495,7,''),(267268,40495,8,''),(267269,40495,9,''),(267270,40495,10,''),(267271,40495,11,''),(267272,40495,12,''),(267273,40495,13,''),(267274,40495,14,''),(267275,40495,15,''),(267276,40495,16,''),(267277,40495,17,''),(267278,40495,18,''),(267279,40495,19,''),(267280,40495,20,''),(267281,40496,1,''),(267282,40496,2,''),(267283,40496,3,''),(267284,40496,4,''),(267285,40496,5,''),(267286,40496,6,''),(267287,40496,7,''),(267288,40496,8,''),(267289,40496,9,''),(267290,40496,10,''),(267291,40496,11,''),(267292,40496,12,''),(267293,40496,13,''),(267294,40496,14,''),(267295,40496,15,''),(267296,40496,16,''),(267297,40496,17,''),(267298,40496,18,''),(267299,40496,19,''),(267300,40496,20,''),(267301,40498,1,''),(267302,40498,2,''),(267303,40498,3,''),(267304,40498,4,''),(267305,40498,5,''),(267306,40498,6,''),(267307,40498,7,''),(267308,40498,8,''),(267309,40498,9,''),(267310,40498,10,''),(267311,40498,11,''),(267312,40498,12,''),(267313,40498,13,''),(267314,40498,14,''),(267315,40498,15,''),(267316,40498,16,''),(267317,40498,17,''),(267318,40498,18,''),(267319,40498,19,''),(267320,40498,20,''),(267321,40499,1,''),(267322,40499,2,''),(267323,40499,3,''),(267324,40499,4,''),(267325,40499,5,''),(267326,40499,6,''),(267327,40499,7,''),(267328,40499,8,''),(267329,40499,9,''),(267330,40499,10,''),(267331,40499,11,''),(267332,40499,12,''),(267333,40499,13,''),(267334,40499,14,''),(267335,40499,15,''),(267336,40499,16,''),(267337,40499,17,''),(267338,40499,18,''),(267339,40499,19,''),(267340,40499,20,''),(267341,40500,1,''),(267342,40500,2,''),(267343,40500,3,''),(267344,40500,4,''),(267345,40500,5,''),(267346,40500,6,''),(267347,40500,7,''),(267348,40500,8,''),(267349,40500,9,''),(267350,40500,10,''),(267351,40500,11,''),(267352,40500,12,''),(267353,40500,13,''),(267354,40500,14,''),(267355,40500,15,''),(267356,40500,16,''),(267357,40500,17,''),(267358,40500,18,''),(267359,40500,19,''),(267360,40500,20,''),(267361,40558,1,''),(267362,40558,2,''),(267363,40558,3,''),(267364,40558,4,''),(267365,40558,5,''),(267366,40558,6,''),(267367,40558,7,''),(267368,40558,8,''),(267369,40558,9,''),(267370,40558,10,''),(267371,40558,11,''),(267372,40558,12,''),(267373,40558,13,''),(267374,40558,14,''),(267375,40558,15,''),(267376,40558,16,''),(267377,40558,17,''),(267378,40558,18,''),(267379,40558,19,''),(267380,40558,20,''),(267381,40580,1,''),(267382,40580,2,''),(267383,40580,3,''),(267384,40580,4,''),(267385,40580,5,''),(267386,40580,6,''),(267387,40580,7,''),(267388,40580,8,''),(267389,40580,9,''),(267390,40580,10,''),(267391,40580,11,''),(267392,40580,12,''),(267393,40580,13,''),(267394,40580,14,''),(267395,40580,15,''),(267396,40580,16,''),(267397,40580,17,''),(267398,40580,18,''),(267399,40580,19,''),(267400,40580,20,''),(267401,40581,1,''),(267402,40581,2,''),(267403,40581,3,''),(267404,40581,4,''),(267405,40581,5,''),(267406,40581,6,''),(267407,40581,7,''),(267408,40581,8,''),(267409,40581,9,''),(267410,40581,10,''),(267411,40581,11,''),(267412,40581,12,''),(267413,40581,13,''),(267414,40581,14,''),(267415,40581,15,''),(267416,40581,16,''),(267417,40581,17,''),(267418,40581,18,''),(267419,40581,19,''),(267420,40581,20,''),(267421,40582,1,''),(267422,40582,2,''),(267423,40582,3,''),(267424,40582,4,''),(267425,40582,5,''),(267426,40582,6,''),(267427,40582,7,''),(267428,40582,8,''),(267429,40582,9,''),(267430,40582,10,''),(267431,40582,11,''),(267432,40582,12,''),(267433,40582,13,''),(267434,40582,14,''),(267435,40582,15,''),(267436,40582,16,''),(267437,40582,17,''),(267438,40582,18,''),(267439,40582,19,''),(267440,40582,20,''),(267441,40651,1,''),(267442,40651,2,''),(267443,40651,3,''),(267444,40651,4,''),(267445,40651,5,''),(267446,40651,6,''),(267447,40651,7,''),(267448,40651,8,''),(267449,40651,9,''),(267450,40651,10,''),(267451,40651,11,''),(267452,40651,12,''),(267453,40651,13,''),(267454,40651,14,''),(267455,40651,15,''),(267456,40651,16,''),(267457,40651,17,''),(267458,40651,18,''),(267459,40651,19,''),(267460,40651,20,''),(267461,40653,1,''),(267462,40653,2,''),(267463,40653,3,''),(267464,40653,4,''),(267465,40653,5,''),(267466,40653,6,''),(267467,40653,7,''),(267468,40653,8,''),(267469,40653,9,''),(267470,40653,10,''),(267471,40653,11,''),(267472,40653,12,''),(267473,40653,13,''),(267474,40653,14,''),(267475,40653,15,''),(267476,40653,16,''),(267477,40653,17,''),(267478,40653,18,''),(267479,40653,19,''),(267480,40653,20,''),(267481,40664,1,''),(267482,40664,2,''),(267483,40664,3,''),(267484,40664,4,''),(267485,40664,5,''),(267486,40664,6,''),(267487,40664,7,''),(267488,40664,8,''),(267489,40664,9,''),(267490,40664,10,''),(267491,40664,11,''),(267492,40664,12,''),(267493,40664,13,''),(267494,40664,14,''),(267495,40664,15,''),(267496,40664,16,''),(267497,40664,17,''),(267498,40664,18,''),(267499,40664,19,''),(267500,40664,20,''),(267501,40665,1,''),(267502,40665,2,''),(267503,40665,3,''),(267504,40665,4,''),(267505,40665,5,''),(267506,40665,6,''),(267507,40665,7,''),(267508,40665,8,''),(267509,40665,9,''),(267510,40665,10,''),(267511,40665,11,''),(267512,40665,12,''),(267513,40665,13,''),(267514,40665,14,''),(267515,40665,15,''),(267516,40665,16,''),(267517,40665,17,''),(267518,40665,18,''),(267519,40665,19,''),(267520,40665,20,''),(267521,40675,1,''),(267522,40675,2,''),(267523,40675,3,''),(267524,40675,4,''),(267525,40675,5,''),(267526,40675,6,''),(267527,40675,7,''),(267528,40675,8,''),(267529,40675,9,''),(267530,40675,10,''),(267531,40675,11,''),(267532,40675,12,''),(267533,40675,13,''),(267534,40675,14,''),(267535,40675,15,''),(267536,40675,16,''),(267537,40675,17,''),(267538,40675,18,''),(267539,40675,19,''),(267540,40675,20,''),(267541,40676,1,''),(267542,40676,2,''),(267543,40676,3,''),(267544,40676,4,''),(267545,40676,5,''),(267546,40676,6,''),(267547,40676,7,''),(267548,40676,8,''),(267549,40676,9,''),(267550,40676,10,''),(267551,40676,11,''),(267552,40676,12,''),(267553,40676,13,''),(267554,40676,14,''),(267555,40676,15,''),(267556,40676,16,''),(267557,40676,17,''),(267558,40676,18,''),(267559,40676,19,''),(267560,40676,20,''),(267561,40678,1,''),(267562,40678,2,''),(267563,40678,3,''),(267564,40678,4,''),(267565,40678,5,''),(267566,40678,6,''),(267567,40678,7,''),(267568,40678,8,''),(267569,40678,9,''),(267570,40678,10,''),(267571,40678,11,''),(267572,40678,12,''),(267573,40678,13,''),(267574,40678,14,''),(267575,40678,15,''),(267576,40678,16,''),(267577,40678,17,''),(267578,40678,18,''),(267579,40678,19,''),(267580,40678,20,''),(267581,40707,1,''),(267582,40707,2,''),(267583,40707,3,''),(267584,40707,4,''),(267585,40707,5,''),(267586,40707,6,''),(267587,40707,7,''),(267588,40707,8,''),(267589,40707,9,''),(267590,40707,10,''),(267591,40707,11,''),(267592,40707,12,''),(267593,40707,13,''),(267594,40707,14,''),(267595,40707,15,''),(267596,40707,16,''),(267597,40707,17,''),(267598,40707,18,''),(267599,40707,19,''),(267600,40707,20,''),(267601,40709,1,''),(267602,40709,2,''),(267603,40709,3,''),(267604,40709,4,''),(267605,40709,5,''),(267606,40709,6,''),(267607,40709,7,''),(267608,40709,8,''),(267609,40709,9,''),(267610,40709,10,''),(267611,40709,11,''),(267612,40709,12,''),(267613,40709,13,''),(267614,40709,14,''),(267615,40709,15,''),(267616,40709,16,''),(267617,40709,17,''),(267618,40709,18,''),(267619,40709,19,''),(267620,40709,20,''),(267621,40733,1,''),(267622,40733,2,''),(267623,40733,3,''),(267624,40733,4,''),(267625,40733,5,''),(267626,40733,6,''),(267627,40733,7,''),(267628,40733,8,''),(267629,40733,9,''),(267630,40733,10,''),(267631,40733,11,''),(267632,40733,12,''),(267633,40733,13,''),(267634,40733,14,''),(267635,40733,15,''),(267636,40733,16,''),(267637,40733,17,''),(267638,40733,18,''),(267639,40733,19,''),(267640,40733,20,''),(267641,40736,1,''),(267642,40736,2,''),(267643,40736,3,''),(267644,40736,4,''),(267645,40736,5,''),(267646,40736,6,''),(267647,40736,7,''),(267648,40736,8,''),(267649,40736,9,''),(267650,40736,10,''),(267651,40736,11,''),(267652,40736,12,''),(267653,40736,13,''),(267654,40736,14,''),(267655,40736,15,''),(267656,40736,16,''),(267657,40736,17,''),(267658,40736,18,''),(267659,40736,19,''),(267660,40736,20,''),(267661,40738,1,''),(267662,40738,2,''),(267663,40738,3,''),(267664,40738,4,''),(267665,40738,5,''),(267666,40738,6,''),(267667,40738,7,''),(267668,40738,8,''),(267669,40738,9,''),(267670,40738,10,''),(267671,40738,11,''),(267672,40738,12,''),(267673,40738,13,''),(267674,40738,14,''),(267675,40738,15,''),(267676,40738,16,''),(267677,40738,17,''),(267678,40738,18,''),(267679,40738,19,''),(267680,40738,20,''),(267681,40741,1,''),(267682,40741,2,''),(267683,40741,3,''),(267684,40741,4,''),(267685,40741,5,''),(267686,40741,6,''),(267687,40741,7,''),(267688,40741,8,''),(267689,40741,9,''),(267690,40741,10,''),(267691,40741,11,''),(267692,40741,12,''),(267693,40741,13,''),(267694,40741,14,''),(267695,40741,15,''),(267696,40741,16,''),(267697,40741,17,''),(267698,40741,18,''),(267699,40741,19,''),(267700,40741,20,''),(267701,40802,1,''),(267702,40802,2,''),(267703,40802,3,''),(267704,40802,4,''),(267705,40802,5,''),(267706,40802,6,''),(267707,40802,7,''),(267708,40802,8,''),(267709,40802,9,''),(267710,40802,10,''),(267711,40802,11,''),(267712,40802,12,''),(267713,40802,13,''),(267714,40802,14,''),(267715,40802,15,''),(267716,40802,16,''),(267717,40802,17,''),(267718,40802,18,''),(267719,40802,19,''),(267720,40802,20,''),(267721,40817,1,''),(267722,40817,2,''),(267723,40817,3,''),(267724,40817,4,''),(267725,40817,5,''),(267726,40817,6,''),(267727,40817,7,''),(267728,40817,8,''),(267729,40817,9,''),(267730,40817,10,''),(267731,40817,11,''),(267732,40817,12,''),(267733,40817,13,''),(267734,40817,14,''),(267735,40817,15,''),(267736,40817,16,''),(267737,40817,17,''),(267738,40817,18,''),(267739,40817,19,''),(267740,40817,20,''),(267741,40877,1,''),(267742,40877,2,''),(267743,40877,3,''),(267744,40877,4,''),(267745,40877,5,''),(267746,40877,6,''),(267747,40877,7,''),(267748,40877,8,''),(267749,40877,9,''),(267750,40877,10,''),(267751,40877,11,''),(267752,40877,12,''),(267753,40877,13,''),(267754,40877,14,''),(267755,40877,15,''),(267756,40877,16,''),(267757,40877,17,''),(267758,40877,18,''),(267759,40877,19,''),(267760,40877,20,''),(267761,40886,1,''),(267762,40886,2,''),(267763,40886,3,''),(267764,40886,4,''),(267765,40886,5,''),(267766,40886,6,''),(267767,40886,7,''),(267768,40886,8,''),(267769,40886,9,''),(267770,40886,10,''),(267771,40886,11,''),(267772,40886,12,''),(267773,40886,13,''),(267774,40886,14,''),(267775,40886,15,''),(267776,40886,16,''),(267777,40886,17,''),(267778,40886,18,''),(267779,40886,19,''),(267780,40886,20,''),(267781,41039,1,''),(267782,41039,2,''),(267783,41039,3,''),(267784,41039,4,''),(267785,41039,5,''),(267786,41039,6,''),(267787,41039,7,''),(267788,41039,8,''),(267789,41039,9,''),(267790,41039,10,''),(267791,41039,11,''),(267792,41039,12,''),(267793,41039,13,''),(267794,41039,14,''),(267795,41039,15,''),(267796,41039,16,''),(267797,41039,17,''),(267798,41039,18,''),(267799,41039,19,''),(267800,41039,20,''),(267801,41062,1,''),(267802,41062,2,''),(267803,41062,3,''),(267804,41062,4,''),(267805,41062,5,''),(267806,41062,6,''),(267807,41062,7,''),(267808,41062,8,''),(267809,41062,9,''),(267810,41062,10,''),(267811,41062,11,''),(267812,41062,12,''),(267813,41062,13,''),(267814,41062,14,''),(267815,41062,15,''),(267816,41062,16,''),(267817,41062,17,''),(267818,41062,18,''),(267819,41062,19,''),(267820,41062,20,''),(267821,41122,1,''),(267822,41122,2,''),(267823,41122,3,''),(267824,41122,4,''),(267825,41122,5,''),(267826,41122,6,''),(267827,41122,7,''),(267828,41122,8,''),(267829,41122,9,''),(267830,41122,10,''),(267831,41122,11,''),(267832,41122,12,''),(267833,41122,13,''),(267834,41122,14,''),(267835,41122,15,''),(267836,41122,16,''),(267837,41122,17,''),(267838,41122,18,''),(267839,41122,19,''),(267840,41122,20,''),(267841,41156,1,''),(267842,41156,2,''),(267843,41156,3,''),(267844,41156,4,''),(267845,41156,5,''),(267846,41156,6,''),(267847,41156,7,''),(267848,41156,8,''),(267849,41156,9,''),(267850,41156,10,''),(267851,41156,11,''),(267852,41156,12,''),(267853,41156,13,''),(267854,41156,14,''),(267855,41156,15,''),(267856,41156,16,''),(267857,41156,17,''),(267858,41156,18,''),(267859,41156,19,''),(267860,41156,20,''),(267861,41168,1,''),(267862,41168,2,''),(267863,41168,3,''),(267864,41168,4,''),(267865,41168,5,''),(267866,41168,6,''),(267867,41168,7,''),(267868,41168,8,''),(267869,41168,9,''),(267870,41168,10,''),(267871,41168,11,''),(267872,41168,12,''),(267873,41168,13,''),(267874,41168,14,''),(267875,41168,15,''),(267876,41168,16,''),(267877,41168,17,''),(267878,41168,18,''),(267879,41168,19,''),(267880,41168,20,''),(267881,41169,1,''),(267882,41169,2,''),(267883,41169,3,''),(267884,41169,4,''),(267885,41169,5,''),(267886,41169,6,''),(267887,41169,7,''),(267888,41169,8,''),(267889,41169,9,''),(267890,41169,10,''),(267891,41169,11,''),(267892,41169,12,''),(267893,41169,13,''),(267894,41169,14,''),(267895,41169,15,''),(267896,41169,16,''),(267897,41169,17,''),(267898,41169,18,''),(267899,41169,19,''),(267900,41169,20,''),(267901,41170,1,''),(267902,41170,2,''),(267903,41170,3,''),(267904,41170,4,''),(267905,41170,5,''),(267906,41170,6,''),(267907,41170,7,''),(267908,41170,8,''),(267909,41170,9,''),(267910,41170,10,''),(267911,41170,11,''),(267912,41170,12,''),(267913,41170,13,''),(267914,41170,14,''),(267915,41170,15,''),(267916,41170,16,''),(267917,41170,17,''),(267918,41170,18,''),(267919,41170,19,''),(267920,41170,20,''),(267921,41196,1,''),(267922,41196,2,''),(267923,41196,3,''),(267924,41196,4,''),(267925,41196,5,''),(267926,41196,6,''),(267927,41196,7,''),(267928,41196,8,''),(267929,41196,9,''),(267930,41196,10,''),(267931,41196,11,''),(267932,41196,12,''),(267933,41196,13,''),(267934,41196,14,''),(267935,41196,15,''),(267936,41196,16,''),(267937,41196,17,''),(267938,41196,18,''),(267939,41196,19,''),(267940,41196,20,''),(267941,41198,1,''),(267942,41198,2,''),(267943,41198,3,''),(267944,41198,4,''),(267945,41198,5,''),(267946,41198,6,''),(267947,41198,7,''),(267948,41198,8,''),(267949,41198,9,''),(267950,41198,10,''),(267951,41198,11,''),(267952,41198,12,''),(267953,41198,13,''),(267954,41198,14,''),(267955,41198,15,''),(267956,41198,16,''),(267957,41198,17,''),(267958,41198,18,''),(267959,41198,19,''),(267960,41198,20,''),(267961,41206,1,''),(267962,41206,2,''),(267963,41206,3,''),(267964,41206,4,''),(267965,41206,5,''),(267966,41206,6,''),(267967,41206,7,''),(267968,41206,8,''),(267969,41206,9,''),(267970,41206,10,''),(267971,41206,11,''),(267972,41206,12,''),(267973,41206,13,''),(267974,41206,14,''),(267975,41206,15,''),(267976,41206,16,''),(267977,41206,17,''),(267978,41206,18,''),(267979,41206,19,''),(267980,41206,20,''),(267981,41224,1,''),(267982,41224,2,''),(267983,41224,3,''),(267984,41224,4,''),(267985,41224,5,''),(267986,41224,6,''),(267987,41224,7,''),(267988,41224,8,''),(267989,41224,9,''),(267990,41224,10,''),(267991,41224,11,''),(267992,41224,12,''),(267993,41224,13,''),(267994,41224,14,''),(267995,41224,15,''),(267996,41224,16,''),(267997,41224,17,''),(267998,41224,18,''),(267999,41224,19,''),(268000,41224,20,''),(268001,41225,1,''),(268002,41225,2,''),(268003,41225,3,''),(268004,41225,4,''),(268005,41225,5,''),(268006,41225,6,''),(268007,41225,7,''),(268008,41225,8,''),(268009,41225,9,''),(268010,41225,10,''),(268011,41225,11,''),(268012,41225,12,''),(268013,41225,13,''),(268014,41225,14,''),(268015,41225,15,''),(268016,41225,16,''),(268017,41225,17,''),(268018,41225,18,''),(268019,41225,19,''),(268020,41225,20,''),(268021,41226,1,''),(268022,41226,2,''),(268023,41226,3,''),(268024,41226,4,''),(268025,41226,5,''),(268026,41226,6,''),(268027,41226,7,''),(268028,41226,8,''),(268029,41226,9,''),(268030,41226,10,''),(268031,41226,11,''),(268032,41226,12,''),(268033,41226,13,''),(268034,41226,14,''),(268035,41226,15,''),(268036,41226,16,''),(268037,41226,17,''),(268038,41226,18,''),(268039,41226,19,''),(268040,41226,20,''),(268041,41235,1,''),(268042,41235,2,''),(268043,41235,3,''),(268044,41235,4,''),(268045,41235,5,''),(268046,41235,6,''),(268047,41235,7,''),(268048,41235,8,''),(268049,41235,9,''),(268050,41235,10,''),(268051,41235,11,''),(268052,41235,12,''),(268053,41235,13,''),(268054,41235,14,''),(268055,41235,15,''),(268056,41235,16,''),(268057,41235,17,''),(268058,41235,18,''),(268059,41235,19,''),(268060,41235,20,''),(268061,41291,1,''),(268062,41291,2,''),(268063,41291,3,''),(268064,41291,4,''),(268065,41291,5,''),(268066,41291,6,''),(268067,41291,7,''),(268068,41291,8,''),(268069,41291,9,''),(268070,41291,10,''),(268071,41291,11,''),(268072,41291,12,''),(268073,41291,13,''),(268074,41291,14,''),(268075,41291,15,''),(268076,41291,16,''),(268077,41291,17,''),(268078,41291,18,''),(268079,41291,19,''),(268080,41291,20,''),(268081,41297,1,''),(268082,41297,2,''),(268083,41297,3,''),(268084,41297,4,''),(268085,41297,5,''),(268086,41297,6,''),(268087,41297,7,''),(268088,41297,8,''),(268089,41297,9,''),(268090,41297,10,''),(268091,41297,11,''),(268092,41297,12,''),(268093,41297,13,''),(268094,41297,14,''),(268095,41297,15,''),(268096,41297,16,''),(268097,41297,17,''),(268098,41297,18,''),(268099,41297,19,''),(268100,41297,20,''),(268101,41298,1,''),(268102,41298,2,''),(268103,41298,3,''),(268104,41298,4,''),(268105,41298,5,''),(268106,41298,6,''),(268107,41298,7,''),(268108,41298,8,''),(268109,41298,9,''),(268110,41298,10,''),(268111,41298,11,''),(268112,41298,12,''),(268113,41298,13,''),(268114,41298,14,''),(268115,41298,15,''),(268116,41298,16,''),(268117,41298,17,''),(268118,41298,18,''),(268119,41298,19,''),(268120,41298,20,''),(268121,41299,1,''),(268122,41299,2,''),(268123,41299,3,''),(268124,41299,4,''),(268125,41299,5,''),(268126,41299,6,''),(268127,41299,7,''),(268128,41299,8,''),(268129,41299,9,''),(268130,41299,10,''),(268131,41299,11,''),(268132,41299,12,''),(268133,41299,13,''),(268134,41299,14,''),(268135,41299,15,''),(268136,41299,16,''),(268137,41299,17,''),(268138,41299,18,''),(268139,41299,19,''),(268140,41299,20,''),(268141,41300,1,''),(268142,41300,2,''),(268143,41300,3,''),(268144,41300,4,''),(268145,41300,5,''),(268146,41300,6,''),(268147,41300,7,''),(268148,41300,8,''),(268149,41300,9,''),(268150,41300,10,''),(268151,41300,11,''),(268152,41300,12,''),(268153,41300,13,''),(268154,41300,14,''),(268155,41300,15,''),(268156,41300,16,''),(268157,41300,17,''),(268158,41300,18,''),(268159,41300,19,''),(268160,41300,20,''),(268161,41302,1,''),(268162,41302,2,''),(268163,41302,3,''),(268164,41302,4,''),(268165,41302,5,''),(268166,41302,6,''),(268167,41302,7,''),(268168,41302,8,''),(268169,41302,9,''),(268170,41302,10,''),(268171,41302,11,''),(268172,41302,12,''),(268173,41302,13,''),(268174,41302,14,''),(268175,41302,15,''),(268176,41302,16,''),(268177,41302,17,''),(268178,41302,18,''),(268179,41302,19,''),(268180,41302,20,''),(268181,41304,1,''),(268182,41304,2,''),(268183,41304,3,''),(268184,41304,4,''),(268185,41304,5,''),(268186,41304,6,''),(268187,41304,7,''),(268188,41304,8,''),(268189,41304,9,''),(268190,41304,10,''),(268191,41304,11,''),(268192,41304,12,''),(268193,41304,13,''),(268194,41304,14,''),(268195,41304,15,''),(268196,41304,16,''),(268197,41304,17,''),(268198,41304,18,''),(268199,41304,19,''),(268200,41304,20,''),(268201,41305,1,''),(268202,41305,2,''),(268203,41305,3,''),(268204,41305,4,''),(268205,41305,5,''),(268206,41305,6,''),(268207,41305,7,''),(268208,41305,8,''),(268209,41305,9,''),(268210,41305,10,''),(268211,41305,11,''),(268212,41305,12,''),(268213,41305,13,''),(268214,41305,14,''),(268215,41305,15,''),(268216,41305,16,''),(268217,41305,17,''),(268218,41305,18,''),(268219,41305,19,''),(268220,41305,20,''),(268221,41308,1,''),(268222,41308,2,''),(268223,41308,3,''),(268224,41308,4,''),(268225,41308,5,''),(268226,41308,6,''),(268227,41308,7,''),(268228,41308,8,''),(268229,41308,9,''),(268230,41308,10,''),(268231,41308,11,''),(268232,41308,12,''),(268233,41308,13,''),(268234,41308,14,''),(268235,41308,15,''),(268236,41308,16,''),(268237,41308,17,''),(268238,41308,18,''),(268239,41308,19,''),(268240,41308,20,''),(268241,41318,1,''),(268242,41318,2,''),(268243,41318,3,''),(268244,41318,4,''),(268245,41318,5,''),(268246,41318,6,''),(268247,41318,7,''),(268248,41318,8,''),(268249,41318,9,''),(268250,41318,10,''),(268251,41318,11,''),(268252,41318,12,''),(268253,41318,13,''),(268254,41318,14,''),(268255,41318,15,''),(268256,41318,16,''),(268257,41318,17,''),(268258,41318,18,''),(268259,41318,19,''),(268260,41318,20,''),(268261,41326,1,''),(268262,41326,2,''),(268263,41326,3,''),(268264,41326,4,''),(268265,41326,5,''),(268266,41326,6,''),(268267,41326,7,''),(268268,41326,8,''),(268269,41326,9,''),(268270,41326,10,''),(268271,41326,11,''),(268272,41326,12,''),(268273,41326,13,''),(268274,41326,14,''),(268275,41326,15,''),(268276,41326,16,''),(268277,41326,17,''),(268278,41326,18,''),(268279,41326,19,''),(268280,41326,20,''),(268281,41327,1,''),(268282,41327,2,''),(268283,41327,3,''),(268284,41327,4,''),(268285,41327,5,''),(268286,41327,6,''),(268287,41327,7,''),(268288,41327,8,''),(268289,41327,9,''),(268290,41327,10,''),(268291,41327,11,''),(268292,41327,12,''),(268293,41327,13,''),(268294,41327,14,''),(268295,41327,15,''),(268296,41327,16,''),(268297,41327,17,''),(268298,41327,18,''),(268299,41327,19,''),(268300,41327,20,''),(268301,41328,1,''),(268302,41328,2,''),(268303,41328,3,''),(268304,41328,4,''),(268305,41328,5,''),(268306,41328,6,''),(268307,41328,7,''),(268308,41328,8,''),(268309,41328,9,''),(268310,41328,10,''),(268311,41328,11,''),(268312,41328,12,''),(268313,41328,13,''),(268314,41328,14,''),(268315,41328,15,''),(268316,41328,16,''),(268317,41328,17,''),(268318,41328,18,''),(268319,41328,19,''),(268320,41328,20,''),(268321,41329,1,''),(268322,41329,2,''),(268323,41329,3,''),(268324,41329,4,''),(268325,41329,5,''),(268326,41329,6,''),(268327,41329,7,''),(268328,41329,8,''),(268329,41329,9,''),(268330,41329,10,''),(268331,41329,11,''),(268332,41329,12,''),(268333,41329,13,''),(268334,41329,14,''),(268335,41329,15,''),(268336,41329,16,''),(268337,41329,17,''),(268338,41329,18,''),(268339,41329,19,''),(268340,41329,20,''),(268341,41331,1,''),(268342,41331,2,''),(268343,41331,3,''),(268344,41331,4,''),(268345,41331,5,''),(268346,41331,6,''),(268347,41331,7,''),(268348,41331,8,''),(268349,41331,9,''),(268350,41331,10,''),(268351,41331,11,''),(268352,41331,12,''),(268353,41331,13,''),(268354,41331,14,''),(268355,41331,15,''),(268356,41331,16,''),(268357,41331,17,''),(268358,41331,18,''),(268359,41331,19,''),(268360,41331,20,''),(268361,41338,1,''),(268362,41338,2,''),(268363,41338,3,''),(268364,41338,4,''),(268365,41338,5,''),(268366,41338,6,''),(268367,41338,7,''),(268368,41338,8,''),(268369,41338,9,''),(268370,41338,10,''),(268371,41338,11,''),(268372,41338,12,''),(268373,41338,13,''),(268374,41338,14,''),(268375,41338,15,''),(268376,41338,16,''),(268377,41338,17,''),(268378,41338,18,''),(268379,41338,19,''),(268380,41338,20,''),(268381,41345,1,''),(268382,41345,2,''),(268383,41345,3,''),(268384,41345,4,''),(268385,41345,5,''),(268386,41345,6,''),(268387,41345,7,''),(268388,41345,8,''),(268389,41345,9,''),(268390,41345,10,''),(268391,41345,11,''),(268392,41345,12,''),(268393,41345,13,''),(268394,41345,14,''),(268395,41345,15,''),(268396,41345,16,''),(268397,41345,17,''),(268398,41345,18,''),(268399,41345,19,''),(268400,41345,20,''),(268401,41368,1,''),(268402,41368,2,''),(268403,41368,3,''),(268404,41368,4,''),(268405,41368,5,''),(268406,41368,6,''),(268407,41368,7,''),(268408,41368,8,''),(268409,41368,9,''),(268410,41368,10,''),(268411,41368,11,''),(268412,41368,12,''),(268413,41368,13,''),(268414,41368,14,''),(268415,41368,15,''),(268416,41368,16,''),(268417,41368,17,''),(268418,41368,18,''),(268419,41368,19,''),(268420,41368,20,''),(268421,41369,1,''),(268422,41369,2,''),(268423,41369,3,''),(268424,41369,4,''),(268425,41369,5,''),(268426,41369,6,''),(268427,41369,7,''),(268428,41369,8,''),(268429,41369,9,''),(268430,41369,10,''),(268431,41369,11,''),(268432,41369,12,''),(268433,41369,13,''),(268434,41369,14,''),(268435,41369,15,''),(268436,41369,16,''),(268437,41369,17,''),(268438,41369,18,''),(268439,41369,19,''),(268440,41369,20,''),(268441,41389,1,''),(268442,41389,2,''),(268443,41389,3,''),(268444,41389,4,''),(268445,41389,5,''),(268446,41389,6,''),(268447,41389,7,''),(268448,41389,8,''),(268449,41389,9,''),(268450,41389,10,''),(268451,41389,11,''),(268452,41389,12,''),(268453,41389,13,''),(268454,41389,14,''),(268455,41389,15,''),(268456,41389,16,''),(268457,41389,17,''),(268458,41389,18,''),(268459,41389,19,''),(268460,41389,20,''),(268461,41390,1,''),(268462,41390,2,''),(268463,41390,3,''),(268464,41390,4,''),(268465,41390,5,''),(268466,41390,6,''),(268467,41390,7,''),(268468,41390,8,''),(268469,41390,9,''),(268470,41390,10,''),(268471,41390,11,''),(268472,41390,12,''),(268473,41390,13,''),(268474,41390,14,''),(268475,41390,15,''),(268476,41390,16,''),(268477,41390,17,''),(268478,41390,18,''),(268479,41390,19,''),(268480,41390,20,''),(268481,41436,1,''),(268482,41436,2,''),(268483,41436,3,''),(268484,41436,4,''),(268485,41436,5,''),(268486,41436,6,''),(268487,41436,7,''),(268488,41436,8,''),(268489,41436,9,''),(268490,41436,10,''),(268491,41436,11,''),(268492,41436,12,''),(268493,41436,13,''),(268494,41436,14,''),(268495,41436,15,''),(268496,41436,16,''),(268497,41436,17,''),(268498,41436,18,''),(268499,41436,19,''),(268500,41436,20,''),(268501,41497,1,''),(268502,41497,2,''),(268503,41497,3,''),(268504,41497,4,''),(268505,41497,5,''),(268506,41497,6,''),(268507,41497,7,''),(268508,41497,8,''),(268509,41497,9,''),(268510,41497,10,''),(268511,41497,11,''),(268512,41497,12,''),(268513,41497,13,''),(268514,41497,14,''),(268515,41497,15,''),(268516,41497,16,''),(268517,41497,17,''),(268518,41497,18,''),(268519,41497,19,''),(268520,41497,20,''),(268521,41500,1,''),(268522,41500,2,''),(268523,41500,3,''),(268524,41500,4,''),(268525,41500,5,''),(268526,41500,6,''),(268527,41500,7,''),(268528,41500,8,''),(268529,41500,9,''),(268530,41500,10,''),(268531,41500,11,''),(268532,41500,12,''),(268533,41500,13,''),(268534,41500,14,''),(268535,41500,15,''),(268536,41500,16,''),(268537,41500,17,''),(268538,41500,18,''),(268539,41500,19,''),(268540,41500,20,''),(268541,41564,1,''),(268542,41564,2,''),(268543,41564,3,''),(268544,41564,4,''),(268545,41564,5,''),(268546,41564,6,''),(268547,41564,7,''),(268548,41564,8,''),(268549,41564,9,''),(268550,41564,10,''),(268551,41564,11,''),(268552,41564,12,''),(268553,41564,13,''),(268554,41564,14,''),(268555,41564,15,''),(268556,41564,16,''),(268557,41564,17,''),(268558,41564,18,''),(268559,41564,19,''),(268560,41564,20,''),(268561,41566,1,''),(268562,41566,2,''),(268563,41566,3,''),(268564,41566,4,''),(268565,41566,5,''),(268566,41566,6,''),(268567,41566,7,''),(268568,41566,8,''),(268569,41566,9,''),(268570,41566,10,''),(268571,41566,11,''),(268572,41566,12,''),(268573,41566,13,''),(268574,41566,14,''),(268575,41566,15,''),(268576,41566,16,''),(268577,41566,17,''),(268578,41566,18,''),(268579,41566,19,''),(268580,41566,20,''),(268581,41568,1,''),(268582,41568,2,''),(268583,41568,3,''),(268584,41568,4,''),(268585,41568,5,''),(268586,41568,6,''),(268587,41568,7,''),(268588,41568,8,''),(268589,41568,9,''),(268590,41568,10,''),(268591,41568,11,''),(268592,41568,12,''),(268593,41568,13,''),(268594,41568,14,''),(268595,41568,15,''),(268596,41568,16,''),(268597,41568,17,''),(268598,41568,18,''),(268599,41568,19,''),(268600,41568,20,''),(268601,41580,1,''),(268602,41580,2,''),(268603,41580,3,''),(268604,41580,4,''),(268605,41580,5,''),(268606,41580,6,''),(268607,41580,7,''),(268608,41580,8,''),(268609,41580,9,''),(268610,41580,10,''),(268611,41580,11,''),(268612,41580,12,''),(268613,41580,13,''),(268614,41580,14,''),(268615,41580,15,''),(268616,41580,16,''),(268617,41580,17,''),(268618,41580,18,''),(268619,41580,19,''),(268620,41580,20,''),(268621,41597,1,''),(268622,41597,2,''),(268623,41597,3,''),(268624,41597,4,''),(268625,41597,5,''),(268626,41597,6,''),(268627,41597,7,''),(268628,41597,8,''),(268629,41597,9,''),(268630,41597,10,''),(268631,41597,11,''),(268632,41597,12,''),(268633,41597,13,''),(268634,41597,14,''),(268635,41597,15,''),(268636,41597,16,''),(268637,41597,17,''),(268638,41597,18,''),(268639,41597,19,''),(268640,41597,20,''),(268641,41620,1,''),(268642,41620,2,''),(268643,41620,3,''),(268644,41620,4,''),(268645,41620,5,''),(268646,41620,6,''),(268647,41620,7,''),(268648,41620,8,''),(268649,41620,9,''),(268650,41620,10,''),(268651,41620,11,''),(268652,41620,12,''),(268653,41620,13,''),(268654,41620,14,''),(268655,41620,15,''),(268656,41620,16,''),(268657,41620,17,''),(268658,41620,18,''),(268659,41620,19,''),(268660,41620,20,''),(268661,41671,1,''),(268662,41671,2,''),(268663,41671,3,''),(268664,41671,4,''),(268665,41671,5,''),(268666,41671,6,''),(268667,41671,7,''),(268668,41671,8,''),(268669,41671,9,''),(268670,41671,10,''),(268671,41671,11,''),(268672,41671,12,''),(268673,41671,13,''),(268674,41671,14,''),(268675,41671,15,''),(268676,41671,16,''),(268677,41671,17,''),(268678,41671,18,''),(268679,41671,19,''),(268680,41671,20,''),(268681,41672,1,''),(268682,41672,2,''),(268683,41672,3,''),(268684,41672,4,''),(268685,41672,5,''),(268686,41672,6,''),(268687,41672,7,''),(268688,41672,8,''),(268689,41672,9,''),(268690,41672,10,''),(268691,41672,11,''),(268692,41672,12,''),(268693,41672,13,''),(268694,41672,14,''),(268695,41672,15,''),(268696,41672,16,''),(268697,41672,17,''),(268698,41672,18,''),(268699,41672,19,''),(268700,41672,20,''),(268701,41684,1,''),(268702,41684,2,''),(268703,41684,3,''),(268704,41684,4,''),(268705,41684,5,''),(268706,41684,6,''),(268707,41684,7,''),(268708,41684,8,''),(268709,41684,9,''),(268710,41684,10,''),(268711,41684,11,''),(268712,41684,12,''),(268713,41684,13,''),(268714,41684,14,''),(268715,41684,15,''),(268716,41684,16,''),(268717,41684,17,''),(268718,41684,18,''),(268719,41684,19,''),(268720,41684,20,''),(268721,41692,1,''),(268722,41692,2,''),(268723,41692,3,''),(268724,41692,4,''),(268725,41692,5,''),(268726,41692,6,''),(268727,41692,7,''),(268728,41692,8,''),(268729,41692,9,''),(268730,41692,10,''),(268731,41692,11,''),(268732,41692,12,''),(268733,41692,13,''),(268734,41692,14,''),(268735,41692,15,''),(268736,41692,16,''),(268737,41692,17,''),(268738,41692,18,''),(268739,41692,19,''),(268740,41692,20,''),(268741,41693,1,''),(268742,41693,2,''),(268743,41693,3,''),(268744,41693,4,''),(268745,41693,5,''),(268746,41693,6,''),(268747,41693,7,''),(268748,41693,8,''),(268749,41693,9,''),(268750,41693,10,''),(268751,41693,11,''),(268752,41693,12,''),(268753,41693,13,''),(268754,41693,14,''),(268755,41693,15,''),(268756,41693,16,''),(268757,41693,17,''),(268758,41693,18,''),(268759,41693,19,''),(268760,41693,20,''),(268761,41694,1,''),(268762,41694,2,''),(268763,41694,3,''),(268764,41694,4,''),(268765,41694,5,''),(268766,41694,6,''),(268767,41694,7,''),(268768,41694,8,''),(268769,41694,9,''),(268770,41694,10,''),(268771,41694,11,''),(268772,41694,12,''),(268773,41694,13,''),(268774,41694,14,''),(268775,41694,15,''),(268776,41694,16,''),(268777,41694,17,''),(268778,41694,18,''),(268779,41694,19,''),(268780,41694,20,''),(268781,41695,1,''),(268782,41695,2,''),(268783,41695,3,''),(268784,41695,4,''),(268785,41695,5,''),(268786,41695,6,''),(268787,41695,7,''),(268788,41695,8,''),(268789,41695,9,''),(268790,41695,10,''),(268791,41695,11,''),(268792,41695,12,''),(268793,41695,13,''),(268794,41695,14,''),(268795,41695,15,''),(268796,41695,16,''),(268797,41695,17,''),(268798,41695,18,''),(268799,41695,19,''),(268800,41695,20,''),(268801,41696,1,''),(268802,41696,2,''),(268803,41696,3,''),(268804,41696,4,''),(268805,41696,5,''),(268806,41696,6,''),(268807,41696,7,''),(268808,41696,8,''),(268809,41696,9,''),(268810,41696,10,''),(268811,41696,11,''),(268812,41696,12,''),(268813,41696,13,''),(268814,41696,14,''),(268815,41696,15,''),(268816,41696,16,''),(268817,41696,17,''),(268818,41696,18,''),(268819,41696,19,''),(268820,41696,20,''),(268821,41712,1,''),(268822,41712,2,''),(268823,41712,3,''),(268824,41712,4,''),(268825,41712,5,''),(268826,41712,6,''),(268827,41712,7,''),(268828,41712,8,''),(268829,41712,9,''),(268830,41712,10,''),(268831,41712,11,''),(268832,41712,12,''),(268833,41712,13,''),(268834,41712,14,''),(268835,41712,15,''),(268836,41712,16,''),(268837,41712,17,''),(268838,41712,18,''),(268839,41712,19,''),(268840,41712,20,''),(268841,41755,1,''),(268842,41755,2,''),(268843,41755,3,''),(268844,41755,4,''),(268845,41755,5,''),(268846,41755,6,''),(268847,41755,7,''),(268848,41755,8,''),(268849,41755,9,''),(268850,41755,10,''),(268851,41755,11,''),(268852,41755,12,''),(268853,41755,13,''),(268854,41755,14,''),(268855,41755,15,''),(268856,41755,16,''),(268857,41755,17,''),(268858,41755,18,''),(268859,41755,19,''),(268860,41755,20,''),(268861,41756,1,''),(268862,41756,2,''),(268863,41756,3,''),(268864,41756,4,''),(268865,41756,5,''),(268866,41756,6,''),(268867,41756,7,''),(268868,41756,8,''),(268869,41756,9,''),(268870,41756,10,''),(268871,41756,11,''),(268872,41756,12,''),(268873,41756,13,''),(268874,41756,14,''),(268875,41756,15,''),(268876,41756,16,''),(268877,41756,17,''),(268878,41756,18,''),(268879,41756,19,''),(268880,41756,20,''),(268881,41762,1,''),(268882,41762,2,''),(268883,41762,3,''),(268884,41762,4,''),(268885,41762,5,''),(268886,41762,6,''),(268887,41762,7,''),(268888,41762,8,''),(268889,41762,9,''),(268890,41762,10,''),(268891,41762,11,''),(268892,41762,12,''),(268893,41762,13,''),(268894,41762,14,''),(268895,41762,15,''),(268896,41762,16,''),(268897,41762,17,''),(268898,41762,18,''),(268899,41762,19,''),(268900,41762,20,''),(268901,41768,1,''),(268902,41768,2,''),(268903,41768,3,''),(268904,41768,4,''),(268905,41768,5,''),(268906,41768,6,''),(268907,41768,7,''),(268908,41768,8,''),(268909,41768,9,''),(268910,41768,10,''),(268911,41768,11,''),(268912,41768,12,''),(268913,41768,13,''),(268914,41768,14,''),(268915,41768,15,''),(268916,41768,16,''),(268917,41768,17,''),(268918,41768,18,''),(268919,41768,19,''),(268920,41768,20,''),(268921,41836,1,''),(268922,41836,2,''),(268923,41836,3,''),(268924,41836,4,''),(268925,41836,5,''),(268926,41836,6,''),(268927,41836,7,''),(268928,41836,8,''),(268929,41836,9,''),(268930,41836,10,''),(268931,41836,11,''),(268932,41836,12,''),(268933,41836,13,''),(268934,41836,14,''),(268935,41836,15,''),(268936,41836,16,''),(268937,41836,17,''),(268938,41836,18,''),(268939,41836,19,''),(268940,41836,20,''),(268941,41844,1,''),(268942,41844,2,''),(268943,41844,3,''),(268944,41844,4,''),(268945,41844,5,''),(268946,41844,6,''),(268947,41844,7,''),(268948,41844,8,''),(268949,41844,9,''),(268950,41844,10,''),(268951,41844,11,''),(268952,41844,12,''),(268953,41844,13,''),(268954,41844,14,''),(268955,41844,15,''),(268956,41844,16,''),(268957,41844,17,''),(268958,41844,18,''),(268959,41844,19,''),(268960,41844,20,''),(268961,41861,1,''),(268962,41861,2,''),(268963,41861,3,''),(268964,41861,4,''),(268965,41861,5,''),(268966,41861,6,''),(268967,41861,7,''),(268968,41861,8,''),(268969,41861,9,''),(268970,41861,10,''),(268971,41861,11,''),(268972,41861,12,''),(268973,41861,13,''),(268974,41861,14,''),(268975,41861,15,''),(268976,41861,16,''),(268977,41861,17,''),(268978,41861,18,''),(268979,41861,19,''),(268980,41861,20,''),(268981,41880,1,''),(268982,41880,2,''),(268983,41880,3,''),(268984,41880,4,''),(268985,41880,5,''),(268986,41880,6,''),(268987,41880,7,''),(268988,41880,8,''),(268989,41880,9,''),(268990,41880,10,''),(268991,41880,11,''),(268992,41880,12,''),(268993,41880,13,''),(268994,41880,14,''),(268995,41880,15,''),(268996,41880,16,''),(268997,41880,17,''),(268998,41880,18,''),(268999,41880,19,''),(269000,41880,20,''),(269001,41881,1,''),(269002,41881,2,''),(269003,41881,3,''),(269004,41881,4,''),(269005,41881,5,''),(269006,41881,6,''),(269007,41881,7,''),(269008,41881,8,''),(269009,41881,9,''),(269010,41881,10,''),(269011,41881,11,''),(269012,41881,12,''),(269013,41881,13,''),(269014,41881,14,''),(269015,41881,15,''),(269016,41881,16,''),(269017,41881,17,''),(269018,41881,18,''),(269019,41881,19,''),(269020,41881,20,''),(269021,41891,1,''),(269022,41891,2,''),(269023,41891,3,''),(269024,41891,4,''),(269025,41891,5,''),(269026,41891,6,''),(269027,41891,7,''),(269028,41891,8,''),(269029,41891,9,''),(269030,41891,10,''),(269031,41891,11,''),(269032,41891,12,''),(269033,41891,13,''),(269034,41891,14,''),(269035,41891,15,''),(269036,41891,16,''),(269037,41891,17,''),(269038,41891,18,''),(269039,41891,19,''),(269040,41891,20,''),(269041,41893,1,''),(269042,41893,2,''),(269043,41893,3,''),(269044,41893,4,''),(269045,41893,5,''),(269046,41893,6,''),(269047,41893,7,''),(269048,41893,8,''),(269049,41893,9,''),(269050,41893,10,''),(269051,41893,11,''),(269052,41893,12,''),(269053,41893,13,''),(269054,41893,14,''),(269055,41893,15,''),(269056,41893,16,''),(269057,41893,17,''),(269058,41893,18,''),(269059,41893,19,''),(269060,41893,20,''),(269061,42023,1,''),(269062,42023,2,''),(269063,42023,3,''),(269064,42023,4,''),(269065,42023,5,''),(269066,42023,6,''),(269067,42023,7,''),(269068,42023,8,''),(269069,42023,9,''),(269070,42023,10,''),(269071,42023,11,''),(269072,42023,12,''),(269073,42023,13,''),(269074,42023,14,''),(269075,42023,15,''),(269076,42023,16,''),(269077,42023,17,''),(269078,42023,18,''),(269079,42023,19,''),(269080,42023,20,''),(269081,42029,1,''),(269082,42029,2,''),(269083,42029,3,''),(269084,42029,4,''),(269085,42029,5,''),(269086,42029,6,''),(269087,42029,7,''),(269088,42029,8,''),(269089,42029,9,''),(269090,42029,10,''),(269091,42029,11,''),(269092,42029,12,''),(269093,42029,13,''),(269094,42029,14,''),(269095,42029,15,''),(269096,42029,16,''),(269097,42029,17,''),(269098,42029,18,''),(269099,42029,19,''),(269100,42029,20,''),(269101,42030,1,''),(269102,42030,2,''),(269103,42030,3,''),(269104,42030,4,''),(269105,42030,5,''),(269106,42030,6,''),(269107,42030,7,''),(269108,42030,8,''),(269109,42030,9,''),(269110,42030,10,''),(269111,42030,11,''),(269112,42030,12,''),(269113,42030,13,''),(269114,42030,14,''),(269115,42030,15,''),(269116,42030,16,''),(269117,42030,17,''),(269118,42030,18,''),(269119,42030,19,''),(269120,42030,20,''),(269121,42066,1,''),(269122,42066,2,''),(269123,42066,3,''),(269124,42066,4,''),(269125,42066,5,''),(269126,42066,6,''),(269127,42066,7,''),(269128,42066,8,''),(269129,42066,9,''),(269130,42066,10,''),(269131,42066,11,''),(269132,42066,12,''),(269133,42066,13,''),(269134,42066,14,''),(269135,42066,15,''),(269136,42066,16,''),(269137,42066,17,''),(269138,42066,18,''),(269139,42066,19,''),(269140,42066,20,''),(269141,42090,1,''),(269142,42090,2,''),(269143,42090,3,''),(269144,42090,4,''),(269145,42090,5,''),(269146,42090,6,''),(269147,42090,7,''),(269148,42090,8,''),(269149,42090,9,''),(269150,42090,10,''),(269151,42090,11,''),(269152,42090,12,''),(269153,42090,13,''),(269154,42090,14,''),(269155,42090,15,''),(269156,42090,16,''),(269157,42090,17,''),(269158,42090,18,''),(269159,42090,19,''),(269160,42090,20,''),(269161,42091,1,''),(269162,42091,2,''),(269163,42091,3,''),(269164,42091,4,''),(269165,42091,5,''),(269166,42091,6,''),(269167,42091,7,''),(269168,42091,8,''),(269169,42091,9,''),(269170,42091,10,''),(269171,42091,11,''),(269172,42091,12,''),(269173,42091,13,''),(269174,42091,14,''),(269175,42091,15,''),(269176,42091,16,''),(269177,42091,17,''),(269178,42091,18,''),(269179,42091,19,''),(269180,42091,20,''),(269181,42098,1,''),(269182,42098,2,''),(269183,42098,3,''),(269184,42098,4,''),(269185,42098,5,''),(269186,42098,6,''),(269187,42098,7,''),(269188,42098,8,''),(269189,42098,9,''),(269190,42098,10,''),(269191,42098,11,''),(269192,42098,12,''),(269193,42098,13,''),(269194,42098,14,''),(269195,42098,15,''),(269196,42098,16,''),(269197,42098,17,''),(269198,42098,18,''),(269199,42098,19,''),(269200,42098,20,''),(269201,42099,1,''),(269202,42099,2,''),(269203,42099,3,''),(269204,42099,4,''),(269205,42099,5,''),(269206,42099,6,''),(269207,42099,7,''),(269208,42099,8,''),(269209,42099,9,''),(269210,42099,10,''),(269211,42099,11,''),(269212,42099,12,''),(269213,42099,13,''),(269214,42099,14,''),(269215,42099,15,''),(269216,42099,16,''),(269217,42099,17,''),(269218,42099,18,''),(269219,42099,19,''),(269220,42099,20,''),(269221,42108,1,''),(269222,42108,2,''),(269223,42108,3,''),(269224,42108,4,''),(269225,42108,5,''),(269226,42108,6,''),(269227,42108,7,''),(269228,42108,8,''),(269229,42108,9,''),(269230,42108,10,''),(269231,42108,11,''),(269232,42108,12,''),(269233,42108,13,''),(269234,42108,14,''),(269235,42108,15,''),(269236,42108,16,''),(269237,42108,17,''),(269238,42108,18,''),(269239,42108,19,''),(269240,42108,20,''),(269241,42110,1,''),(269242,42110,2,''),(269243,42110,3,''),(269244,42110,4,''),(269245,42110,5,''),(269246,42110,6,''),(269247,42110,7,''),(269248,42110,8,''),(269249,42110,9,''),(269250,42110,10,''),(269251,42110,11,''),(269252,42110,12,''),(269253,42110,13,''),(269254,42110,14,''),(269255,42110,15,''),(269256,42110,16,''),(269257,42110,17,''),(269258,42110,18,''),(269259,42110,19,''),(269260,42110,20,''),(269261,42174,1,''),(269262,42174,2,''),(269263,42174,3,''),(269264,42174,4,''),(269265,42174,5,''),(269266,42174,6,''),(269267,42174,7,''),(269268,42174,8,''),(269269,42174,9,''),(269270,42174,10,''),(269271,42174,11,''),(269272,42174,12,''),(269273,42174,13,''),(269274,42174,14,''),(269275,42174,15,''),(269276,42174,16,''),(269277,42174,17,''),(269278,42174,18,''),(269279,42174,19,''),(269280,42174,20,''),(269281,42175,1,''),(269282,42175,2,''),(269283,42175,3,''),(269284,42175,4,''),(269285,42175,5,''),(269286,42175,6,''),(269287,42175,7,''),(269288,42175,8,''),(269289,42175,9,''),(269290,42175,10,''),(269291,42175,11,''),(269292,42175,12,''),(269293,42175,13,''),(269294,42175,14,''),(269295,42175,15,''),(269296,42175,16,''),(269297,42175,17,''),(269298,42175,18,''),(269299,42175,19,''),(269300,42175,20,''),(269301,42176,1,''),(269302,42176,2,''),(269303,42176,3,''),(269304,42176,4,''),(269305,42176,5,''),(269306,42176,6,''),(269307,42176,7,''),(269308,42176,8,''),(269309,42176,9,''),(269310,42176,10,''),(269311,42176,11,''),(269312,42176,12,''),(269313,42176,13,''),(269314,42176,14,''),(269315,42176,15,''),(269316,42176,16,''),(269317,42176,17,''),(269318,42176,18,''),(269319,42176,19,''),(269320,42176,20,''),(269321,42178,1,''),(269322,42178,2,''),(269323,42178,3,''),(269324,42178,4,''),(269325,42178,5,''),(269326,42178,6,''),(269327,42178,7,''),(269328,42178,8,''),(269329,42178,9,''),(269330,42178,10,''),(269331,42178,11,''),(269332,42178,12,''),(269333,42178,13,''),(269334,42178,14,''),(269335,42178,15,''),(269336,42178,16,''),(269337,42178,17,''),(269338,42178,18,''),(269339,42178,19,''),(269340,42178,20,''),(269341,42180,1,''),(269342,42180,2,''),(269343,42180,3,''),(269344,42180,4,''),(269345,42180,5,''),(269346,42180,6,''),(269347,42180,7,''),(269348,42180,8,''),(269349,42180,9,''),(269350,42180,10,''),(269351,42180,11,''),(269352,42180,12,''),(269353,42180,13,''),(269354,42180,14,''),(269355,42180,15,''),(269356,42180,16,''),(269357,42180,17,''),(269358,42180,18,''),(269359,42180,19,''),(269360,42180,20,''),(269361,42187,1,''),(269362,42187,2,''),(269363,42187,3,''),(269364,42187,4,''),(269365,42187,5,''),(269366,42187,6,''),(269367,42187,7,''),(269368,42187,8,''),(269369,42187,9,''),(269370,42187,10,''),(269371,42187,11,''),(269372,42187,12,''),(269373,42187,13,''),(269374,42187,14,''),(269375,42187,15,''),(269376,42187,16,''),(269377,42187,17,''),(269378,42187,18,''),(269379,42187,19,''),(269380,42187,20,''),(269381,42201,1,''),(269382,42201,2,''),(269383,42201,3,''),(269384,42201,4,''),(269385,42201,5,''),(269386,42201,6,''),(269387,42201,7,''),(269388,42201,8,''),(269389,42201,9,''),(269390,42201,10,''),(269391,42201,11,''),(269392,42201,12,''),(269393,42201,13,''),(269394,42201,14,''),(269395,42201,15,''),(269396,42201,16,''),(269397,42201,17,''),(269398,42201,18,''),(269399,42201,19,''),(269400,42201,20,''),(269401,42217,1,''),(269402,42217,2,''),(269403,42217,3,''),(269404,42217,4,''),(269405,42217,5,''),(269406,42217,6,''),(269407,42217,7,''),(269408,42217,8,''),(269409,42217,9,''),(269410,42217,10,''),(269411,42217,11,''),(269412,42217,12,''),(269413,42217,13,''),(269414,42217,14,''),(269415,42217,15,''),(269416,42217,16,''),(269417,42217,17,''),(269418,42217,18,''),(269419,42217,19,''),(269420,42217,20,''),(269421,42233,1,''),(269422,42233,2,''),(269423,42233,3,''),(269424,42233,4,''),(269425,42233,5,''),(269426,42233,6,''),(269427,42233,7,''),(269428,42233,8,''),(269429,42233,9,''),(269430,42233,10,''),(269431,42233,11,''),(269432,42233,12,''),(269433,42233,13,''),(269434,42233,14,''),(269435,42233,15,''),(269436,42233,16,''),(269437,42233,17,''),(269438,42233,18,''),(269439,42233,19,''),(269440,42233,20,''),(269441,42234,1,''),(269442,42234,2,''),(269443,42234,3,''),(269444,42234,4,''),(269445,42234,5,''),(269446,42234,6,''),(269447,42234,7,''),(269448,42234,8,''),(269449,42234,9,''),(269450,42234,10,''),(269451,42234,11,''),(269452,42234,12,''),(269453,42234,13,''),(269454,42234,14,''),(269455,42234,15,''),(269456,42234,16,''),(269457,42234,17,''),(269458,42234,18,''),(269459,42234,19,''),(269460,42234,20,''),(269461,42237,1,''),(269462,42237,2,''),(269463,42237,3,''),(269464,42237,4,''),(269465,42237,5,''),(269466,42237,6,''),(269467,42237,7,''),(269468,42237,8,''),(269469,42237,9,''),(269470,42237,10,''),(269471,42237,11,''),(269472,42237,12,''),(269473,42237,13,''),(269474,42237,14,''),(269475,42237,15,''),(269476,42237,16,''),(269477,42237,17,''),(269478,42237,18,''),(269479,42237,19,''),(269480,42237,20,''),(269481,42247,1,''),(269482,42247,2,''),(269483,42247,3,''),(269484,42247,4,''),(269485,42247,5,''),(269486,42247,6,''),(269487,42247,7,''),(269488,42247,8,''),(269489,42247,9,''),(269490,42247,10,''),(269491,42247,11,''),(269492,42247,12,''),(269493,42247,13,''),(269494,42247,14,''),(269495,42247,15,''),(269496,42247,16,''),(269497,42247,17,''),(269498,42247,18,''),(269499,42247,19,''),(269500,42247,20,''),(269501,42429,1,''),(269502,42429,2,''),(269503,42429,3,''),(269504,42429,4,''),(269505,42429,5,''),(269506,42429,6,''),(269507,42429,7,''),(269508,42429,8,''),(269509,42429,9,''),(269510,42429,10,''),(269511,42429,11,''),(269512,42429,12,''),(269513,42429,13,''),(269514,42429,14,''),(269515,42429,15,''),(269516,42429,16,''),(269517,42429,17,''),(269518,42429,18,''),(269519,42429,19,''),(269520,42429,20,''),(269521,42443,1,''),(269522,42443,2,''),(269523,42443,3,''),(269524,42443,4,''),(269525,42443,5,''),(269526,42443,6,''),(269527,42443,7,''),(269528,42443,8,''),(269529,42443,9,''),(269530,42443,10,''),(269531,42443,11,''),(269532,42443,12,''),(269533,42443,13,''),(269534,42443,14,''),(269535,42443,15,''),(269536,42443,16,''),(269537,42443,17,''),(269538,42443,18,''),(269539,42443,19,''),(269540,42443,20,''),(269541,42469,1,''),(269542,42469,2,''),(269543,42469,3,''),(269544,42469,4,''),(269545,42469,5,''),(269546,42469,6,''),(269547,42469,7,''),(269548,42469,8,''),(269549,42469,9,''),(269550,42469,10,''),(269551,42469,11,''),(269552,42469,12,''),(269553,42469,13,''),(269554,42469,14,''),(269555,42469,15,''),(269556,42469,16,''),(269557,42469,17,''),(269558,42469,18,''),(269559,42469,19,''),(269560,42469,20,''),(269561,42474,1,''),(269562,42474,2,''),(269563,42474,3,''),(269564,42474,4,''),(269565,42474,5,''),(269566,42474,6,''),(269567,42474,7,''),(269568,42474,8,''),(269569,42474,9,''),(269570,42474,10,''),(269571,42474,11,''),(269572,42474,12,''),(269573,42474,13,''),(269574,42474,14,''),(269575,42474,15,''),(269576,42474,16,''),(269577,42474,17,''),(269578,42474,18,''),(269579,42474,19,''),(269580,42474,20,''),(269581,42475,1,''),(269582,42475,2,''),(269583,42475,3,''),(269584,42475,4,''),(269585,42475,5,''),(269586,42475,6,''),(269587,42475,7,''),(269588,42475,8,''),(269589,42475,9,''),(269590,42475,10,''),(269591,42475,11,''),(269592,42475,12,''),(269593,42475,13,''),(269594,42475,14,''),(269595,42475,15,''),(269596,42475,16,''),(269597,42475,17,''),(269598,42475,18,''),(269599,42475,19,''),(269600,42475,20,''),(269601,42476,1,''),(269602,42476,2,''),(269603,42476,3,''),(269604,42476,4,''),(269605,42476,5,''),(269606,42476,6,''),(269607,42476,7,''),(269608,42476,8,''),(269609,42476,9,''),(269610,42476,10,''),(269611,42476,11,''),(269612,42476,12,''),(269613,42476,13,''),(269614,42476,14,''),(269615,42476,15,''),(269616,42476,16,''),(269617,42476,17,''),(269618,42476,18,''),(269619,42476,19,''),(269620,42476,20,''),(269621,42492,1,''),(269622,42492,2,''),(269623,42492,3,''),(269624,42492,4,''),(269625,42492,5,''),(269626,42492,6,''),(269627,42492,7,''),(269628,42492,8,''),(269629,42492,9,''),(269630,42492,10,''),(269631,42492,11,''),(269632,42492,12,''),(269633,42492,13,''),(269634,42492,14,''),(269635,42492,15,''),(269636,42492,16,''),(269637,42492,17,''),(269638,42492,18,''),(269639,42492,19,''),(269640,42492,20,''),(269641,42494,1,''),(269642,42494,2,''),(269643,42494,3,''),(269644,42494,4,''),(269645,42494,5,''),(269646,42494,6,''),(269647,42494,7,''),(269648,42494,8,''),(269649,42494,9,''),(269650,42494,10,''),(269651,42494,11,''),(269652,42494,12,''),(269653,42494,13,''),(269654,42494,14,''),(269655,42494,15,''),(269656,42494,16,''),(269657,42494,17,''),(269658,42494,18,''),(269659,42494,19,''),(269660,42494,20,''),(269661,42495,1,''),(269662,42495,2,''),(269663,42495,3,''),(269664,42495,4,''),(269665,42495,5,''),(269666,42495,6,''),(269667,42495,7,''),(269668,42495,8,''),(269669,42495,9,''),(269670,42495,10,''),(269671,42495,11,''),(269672,42495,12,''),(269673,42495,13,''),(269674,42495,14,''),(269675,42495,15,''),(269676,42495,16,''),(269677,42495,17,''),(269678,42495,18,''),(269679,42495,19,''),(269680,42495,20,''),(269681,42523,1,''),(269682,42523,2,''),(269683,42523,3,''),(269684,42523,4,''),(269685,42523,5,''),(269686,42523,6,''),(269687,42523,7,''),(269688,42523,8,''),(269689,42523,9,''),(269690,42523,10,''),(269691,42523,11,''),(269692,42523,12,''),(269693,42523,13,''),(269694,42523,14,''),(269695,42523,15,''),(269696,42523,16,''),(269697,42523,17,''),(269698,42523,18,''),(269699,42523,19,''),(269700,42523,20,''),(269701,42524,1,''),(269702,42524,2,''),(269703,42524,3,''),(269704,42524,4,''),(269705,42524,5,''),(269706,42524,6,''),(269707,42524,7,''),(269708,42524,8,''),(269709,42524,9,''),(269710,42524,10,''),(269711,42524,11,''),(269712,42524,12,''),(269713,42524,13,''),(269714,42524,14,''),(269715,42524,15,''),(269716,42524,16,''),(269717,42524,17,''),(269718,42524,18,''),(269719,42524,19,''),(269720,42524,20,''),(269721,42567,1,''),(269722,42567,2,''),(269723,42567,3,''),(269724,42567,4,''),(269725,42567,5,''),(269726,42567,6,''),(269727,42567,7,''),(269728,42567,8,''),(269729,42567,9,''),(269730,42567,10,''),(269731,42567,11,''),(269732,42567,12,''),(269733,42567,13,''),(269734,42567,14,''),(269735,42567,15,''),(269736,42567,16,''),(269737,42567,17,''),(269738,42567,18,''),(269739,42567,19,''),(269740,42567,20,''),(269741,42568,1,''),(269742,42568,2,''),(269743,42568,3,''),(269744,42568,4,''),(269745,42568,5,''),(269746,42568,6,''),(269747,42568,7,''),(269748,42568,8,''),(269749,42568,9,''),(269750,42568,10,''),(269751,42568,11,''),(269752,42568,12,''),(269753,42568,13,''),(269754,42568,14,''),(269755,42568,15,''),(269756,42568,16,''),(269757,42568,17,''),(269758,42568,18,''),(269759,42568,19,''),(269760,42568,20,''),(269761,42569,1,''),(269762,42569,2,''),(269763,42569,3,''),(269764,42569,4,''),(269765,42569,5,''),(269766,42569,6,''),(269767,42569,7,''),(269768,42569,8,''),(269769,42569,9,''),(269770,42569,10,''),(269771,42569,11,''),(269772,42569,12,''),(269773,42569,13,''),(269774,42569,14,''),(269775,42569,15,''),(269776,42569,16,''),(269777,42569,17,''),(269778,42569,18,''),(269779,42569,19,''),(269780,42569,20,''),(269781,42570,1,''),(269782,42570,2,''),(269783,42570,3,''),(269784,42570,4,''),(269785,42570,5,''),(269786,42570,6,''),(269787,42570,7,''),(269788,42570,8,''),(269789,42570,9,''),(269790,42570,10,''),(269791,42570,11,''),(269792,42570,12,''),(269793,42570,13,''),(269794,42570,14,''),(269795,42570,15,''),(269796,42570,16,''),(269797,42570,17,''),(269798,42570,18,''),(269799,42570,19,''),(269800,42570,20,''),(269801,42571,1,''),(269802,42571,2,''),(269803,42571,3,''),(269804,42571,4,''),(269805,42571,5,''),(269806,42571,6,''),(269807,42571,7,''),(269808,42571,8,''),(269809,42571,9,''),(269810,42571,10,''),(269811,42571,11,''),(269812,42571,12,''),(269813,42571,13,''),(269814,42571,14,''),(269815,42571,15,''),(269816,42571,16,''),(269817,42571,17,''),(269818,42571,18,''),(269819,42571,19,''),(269820,42571,20,''),(269821,42572,1,''),(269822,42572,2,''),(269823,42572,3,''),(269824,42572,4,''),(269825,42572,5,''),(269826,42572,6,''),(269827,42572,7,''),(269828,42572,8,''),(269829,42572,9,''),(269830,42572,10,''),(269831,42572,11,''),(269832,42572,12,''),(269833,42572,13,''),(269834,42572,14,''),(269835,42572,15,''),(269836,42572,16,''),(269837,42572,17,''),(269838,42572,18,''),(269839,42572,19,''),(269840,42572,20,''),(269841,42573,1,''),(269842,42573,2,''),(269843,42573,3,''),(269844,42573,4,''),(269845,42573,5,''),(269846,42573,6,''),(269847,42573,7,''),(269848,42573,8,''),(269849,42573,9,''),(269850,42573,10,''),(269851,42573,11,''),(269852,42573,12,''),(269853,42573,13,''),(269854,42573,14,''),(269855,42573,15,''),(269856,42573,16,''),(269857,42573,17,''),(269858,42573,18,''),(269859,42573,19,''),(269860,42573,20,''),(269861,42574,1,''),(269862,42574,2,''),(269863,42574,3,''),(269864,42574,4,''),(269865,42574,5,''),(269866,42574,6,''),(269867,42574,7,''),(269868,42574,8,''),(269869,42574,9,''),(269870,42574,10,''),(269871,42574,11,''),(269872,42574,12,''),(269873,42574,13,''),(269874,42574,14,''),(269875,42574,15,''),(269876,42574,16,''),(269877,42574,17,''),(269878,42574,18,''),(269879,42574,19,''),(269880,42574,20,''),(269881,42575,1,''),(269882,42575,2,''),(269883,42575,3,''),(269884,42575,4,''),(269885,42575,5,''),(269886,42575,6,''),(269887,42575,7,''),(269888,42575,8,''),(269889,42575,9,''),(269890,42575,10,''),(269891,42575,11,''),(269892,42575,12,''),(269893,42575,13,''),(269894,42575,14,''),(269895,42575,15,''),(269896,42575,16,''),(269897,42575,17,''),(269898,42575,18,''),(269899,42575,19,''),(269900,42575,20,''),(269901,42576,1,''),(269902,42576,2,''),(269903,42576,3,''),(269904,42576,4,''),(269905,42576,5,''),(269906,42576,6,''),(269907,42576,7,''),(269908,42576,8,''),(269909,42576,9,''),(269910,42576,10,''),(269911,42576,11,''),(269912,42576,12,''),(269913,42576,13,''),(269914,42576,14,''),(269915,42576,15,''),(269916,42576,16,''),(269917,42576,17,''),(269918,42576,18,''),(269919,42576,19,''),(269920,42576,20,''),(269921,42577,1,''),(269922,42577,2,''),(269923,42577,3,''),(269924,42577,4,''),(269925,42577,5,''),(269926,42577,6,''),(269927,42577,7,''),(269928,42577,8,''),(269929,42577,9,''),(269930,42577,10,''),(269931,42577,11,''),(269932,42577,12,''),(269933,42577,13,''),(269934,42577,14,''),(269935,42577,15,''),(269936,42577,16,''),(269937,42577,17,''),(269938,42577,18,''),(269939,42577,19,''),(269940,42577,20,''),(269941,42620,1,''),(269942,42620,2,''),(269943,42620,3,''),(269944,42620,4,''),(269945,42620,5,''),(269946,42620,6,''),(269947,42620,7,''),(269948,42620,8,''),(269949,42620,9,''),(269950,42620,10,''),(269951,42620,11,''),(269952,42620,12,''),(269953,42620,13,''),(269954,42620,14,''),(269955,42620,15,''),(269956,42620,16,''),(269957,42620,17,''),(269958,42620,18,''),(269959,42620,19,''),(269960,42620,20,''),(269961,42706,1,''),(269962,42706,2,''),(269963,42706,3,''),(269964,42706,4,''),(269965,42706,5,''),(269966,42706,6,''),(269967,42706,7,''),(269968,42706,8,''),(269969,42706,9,''),(269970,42706,10,''),(269971,42706,11,''),(269972,42706,12,''),(269973,42706,13,''),(269974,42706,14,''),(269975,42706,15,''),(269976,42706,16,''),(269977,42706,17,''),(269978,42706,18,''),(269979,42706,19,''),(269980,42706,20,''),(269981,42713,1,''),(269982,42713,2,''),(269983,42713,3,''),(269984,42713,4,''),(269985,42713,5,''),(269986,42713,6,''),(269987,42713,7,''),(269988,42713,8,''),(269989,42713,9,''),(269990,42713,10,''),(269991,42713,11,''),(269992,42713,12,''),(269993,42713,13,''),(269994,42713,14,''),(269995,42713,15,''),(269996,42713,16,''),(269997,42713,17,''),(269998,42713,18,''),(269999,42713,19,''),(270000,42713,20,''),(270001,42745,1,''),(270002,42745,2,''),(270003,42745,3,''),(270004,42745,4,''),(270005,42745,5,''),(270006,42745,6,''),(270007,42745,7,''),(270008,42745,8,''),(270009,42745,9,''),(270010,42745,10,''),(270011,42745,11,''),(270012,42745,12,''),(270013,42745,13,''),(270014,42745,14,''),(270015,42745,15,''),(270016,42745,16,''),(270017,42745,17,''),(270018,42745,18,''),(270019,42745,19,''),(270020,42745,20,''),(270021,42821,1,''),(270022,42821,2,''),(270023,42821,3,''),(270024,42821,4,''),(270025,42821,5,''),(270026,42821,6,''),(270027,42821,7,''),(270028,42821,8,''),(270029,42821,9,''),(270030,42821,10,''),(270031,42821,11,''),(270032,42821,12,''),(270033,42821,13,''),(270034,42821,14,''),(270035,42821,15,''),(270036,42821,16,''),(270037,42821,17,''),(270038,42821,18,''),(270039,42821,19,''),(270040,42821,20,''),(270041,42858,1,''),(270042,42858,2,''),(270043,42858,3,''),(270044,42858,4,''),(270045,42858,5,''),(270046,42858,6,''),(270047,42858,7,''),(270048,42858,8,''),(270049,42858,9,''),(270050,42858,10,''),(270051,42858,11,''),(270052,42858,12,''),(270053,42858,13,''),(270054,42858,14,''),(270055,42858,15,''),(270056,42858,16,''),(270057,42858,17,''),(270058,42858,18,''),(270059,42858,19,''),(270060,42858,20,''),(270061,42860,1,''),(270062,42860,2,''),(270063,42860,3,''),(270064,42860,4,''),(270065,42860,5,''),(270066,42860,6,''),(270067,42860,7,''),(270068,42860,8,''),(270069,42860,9,''),(270070,42860,10,''),(270071,42860,11,''),(270072,42860,12,''),(270073,42860,13,''),(270074,42860,14,''),(270075,42860,15,''),(270076,42860,16,''),(270077,42860,17,''),(270078,42860,18,''),(270079,42860,19,''),(270080,42860,20,''),(270081,42868,1,''),(270082,42868,2,''),(270083,42868,3,''),(270084,42868,4,''),(270085,42868,5,''),(270086,42868,6,''),(270087,42868,7,''),(270088,42868,8,''),(270089,42868,9,''),(270090,42868,10,''),(270091,42868,11,''),(270092,42868,12,''),(270093,42868,13,''),(270094,42868,14,''),(270095,42868,15,''),(270096,42868,16,''),(270097,42868,17,''),(270098,42868,18,''),(270099,42868,19,''),(270100,42868,20,''),(270101,43036,1,''),(270102,43036,2,''),(270103,43036,3,''),(270104,43036,4,''),(270105,43036,5,''),(270106,43036,6,''),(270107,43036,7,''),(270108,43036,8,''),(270109,43036,9,''),(270110,43036,10,''),(270111,43036,11,''),(270112,43036,12,''),(270113,43036,13,''),(270114,43036,14,''),(270115,43036,15,''),(270116,43036,16,''),(270117,43036,17,''),(270118,43036,18,''),(270119,43036,19,''),(270120,43036,20,''),(270121,43092,1,''),(270122,43092,2,''),(270123,43092,3,''),(270124,43092,4,''),(270125,43092,5,''),(270126,43092,6,''),(270127,43092,7,''),(270128,43092,8,''),(270129,43092,9,''),(270130,43092,10,''),(270131,43092,11,''),(270132,43092,12,''),(270133,43092,13,''),(270134,43092,14,''),(270135,43092,15,''),(270136,43092,16,''),(270137,43092,17,''),(270138,43092,18,''),(270139,43092,19,''),(270140,43092,20,''),(270141,43093,1,''),(270142,43093,2,''),(270143,43093,3,''),(270144,43093,4,''),(270145,43093,5,''),(270146,43093,6,''),(270147,43093,7,''),(270148,43093,8,''),(270149,43093,9,''),(270150,43093,10,''),(270151,43093,11,''),(270152,43093,12,''),(270153,43093,13,''),(270154,43093,14,''),(270155,43093,15,''),(270156,43093,16,''),(270157,43093,17,''),(270158,43093,18,''),(270159,43093,19,''),(270160,43093,20,''),(270161,43094,1,''),(270162,43094,2,''),(270163,43094,3,''),(270164,43094,4,''),(270165,43094,5,''),(270166,43094,6,''),(270167,43094,7,''),(270168,43094,8,''),(270169,43094,9,''),(270170,43094,10,''),(270171,43094,11,''),(270172,43094,12,''),(270173,43094,13,''),(270174,43094,14,''),(270175,43094,15,''),(270176,43094,16,''),(270177,43094,17,''),(270178,43094,18,''),(270179,43094,19,''),(270180,43094,20,''),(270181,43095,1,''),(270182,43095,2,''),(270183,43095,3,''),(270184,43095,4,''),(270185,43095,5,''),(270186,43095,6,''),(270187,43095,7,''),(270188,43095,8,''),(270189,43095,9,''),(270190,43095,10,''),(270191,43095,11,''),(270192,43095,12,''),(270193,43095,13,''),(270194,43095,14,''),(270195,43095,15,''),(270196,43095,16,''),(270197,43095,17,''),(270198,43095,18,''),(270199,43095,19,''),(270200,43095,20,''),(270201,43096,1,''),(270202,43096,2,''),(270203,43096,3,''),(270204,43096,4,''),(270205,43096,5,''),(270206,43096,6,''),(270207,43096,7,''),(270208,43096,8,''),(270209,43096,9,''),(270210,43096,10,''),(270211,43096,11,''),(270212,43096,12,''),(270213,43096,13,''),(270214,43096,14,''),(270215,43096,15,''),(270216,43096,16,''),(270217,43096,17,''),(270218,43096,18,''),(270219,43096,19,''),(270220,43096,20,''),(270221,43097,1,''),(270222,43097,2,''),(270223,43097,3,''),(270224,43097,4,''),(270225,43097,5,''),(270226,43097,6,''),(270227,43097,7,''),(270228,43097,8,''),(270229,43097,9,''),(270230,43097,10,''),(270231,43097,11,''),(270232,43097,12,''),(270233,43097,13,''),(270234,43097,14,''),(270235,43097,15,''),(270236,43097,16,''),(270237,43097,17,''),(270238,43097,18,''),(270239,43097,19,''),(270240,43097,20,''),(270241,43098,1,''),(270242,43098,2,''),(270243,43098,3,''),(270244,43098,4,''),(270245,43098,5,''),(270246,43098,6,''),(270247,43098,7,''),(270248,43098,8,''),(270249,43098,9,''),(270250,43098,10,''),(270251,43098,11,''),(270252,43098,12,''),(270253,43098,13,''),(270254,43098,14,''),(270255,43098,15,''),(270256,43098,16,''),(270257,43098,17,''),(270258,43098,18,''),(270259,43098,19,''),(270260,43098,20,''),(270261,43099,1,''),(270262,43099,2,''),(270263,43099,3,''),(270264,43099,4,''),(270265,43099,5,''),(270266,43099,6,''),(270267,43099,7,''),(270268,43099,8,''),(270269,43099,9,''),(270270,43099,10,''),(270271,43099,11,''),(270272,43099,12,''),(270273,43099,13,''),(270274,43099,14,''),(270275,43099,15,''),(270276,43099,16,''),(270277,43099,17,''),(270278,43099,18,''),(270279,43099,19,''),(270280,43099,20,''),(270281,43100,1,''),(270282,43100,2,''),(270283,43100,3,''),(270284,43100,4,''),(270285,43100,5,''),(270286,43100,6,''),(270287,43100,7,''),(270288,43100,8,''),(270289,43100,9,''),(270290,43100,10,''),(270291,43100,11,''),(270292,43100,12,''),(270293,43100,13,''),(270294,43100,14,''),(270295,43100,15,''),(270296,43100,16,''),(270297,43100,17,''),(270298,43100,18,''),(270299,43100,19,''),(270300,43100,20,''),(270301,43114,1,''),(270302,43114,2,''),(270303,43114,3,''),(270304,43114,4,''),(270305,43114,5,''),(270306,43114,6,''),(270307,43114,7,''),(270308,43114,8,''),(270309,43114,9,''),(270310,43114,10,''),(270311,43114,11,''),(270312,43114,12,''),(270313,43114,13,''),(270314,43114,14,''),(270315,43114,15,''),(270316,43114,16,''),(270317,43114,17,''),(270318,43114,18,''),(270319,43114,19,''),(270320,43114,20,''),(270321,43132,1,''),(270322,43132,2,''),(270323,43132,3,''),(270324,43132,4,''),(270325,43132,5,''),(270326,43132,6,''),(270327,43132,7,''),(270328,43132,8,''),(270329,43132,9,''),(270330,43132,10,''),(270331,43132,11,''),(270332,43132,12,''),(270333,43132,13,''),(270334,43132,14,''),(270335,43132,15,''),(270336,43132,16,''),(270337,43132,17,''),(270338,43132,18,''),(270339,43132,19,''),(270340,43132,20,''),(270341,43143,1,''),(270342,43143,2,''),(270343,43143,3,''),(270344,43143,4,''),(270345,43143,5,''),(270346,43143,6,''),(270347,43143,7,''),(270348,43143,8,''),(270349,43143,9,''),(270350,43143,10,''),(270351,43143,11,''),(270352,43143,12,''),(270353,43143,13,''),(270354,43143,14,''),(270355,43143,15,''),(270356,43143,16,''),(270357,43143,17,''),(270358,43143,18,''),(270359,43143,19,''),(270360,43143,20,''),(270361,43148,1,''),(270362,43148,2,''),(270363,43148,3,''),(270364,43148,4,''),(270365,43148,5,''),(270366,43148,6,''),(270367,43148,7,''),(270368,43148,8,''),(270369,43148,9,''),(270370,43148,10,''),(270371,43148,11,''),(270372,43148,12,''),(270373,43148,13,''),(270374,43148,14,''),(270375,43148,15,''),(270376,43148,16,''),(270377,43148,17,''),(270378,43148,18,''),(270379,43148,19,''),(270380,43148,20,''),(270381,43176,1,''),(270382,43176,2,''),(270383,43176,3,''),(270384,43176,4,''),(270385,43176,5,''),(270386,43176,6,''),(270387,43176,7,''),(270388,43176,8,''),(270389,43176,9,''),(270390,43176,10,''),(270391,43176,11,''),(270392,43176,12,''),(270393,43176,13,''),(270394,43176,14,''),(270395,43176,15,''),(270396,43176,16,''),(270397,43176,17,''),(270398,43176,18,''),(270399,43176,19,''),(270400,43176,20,''),(270401,43329,1,''),(270402,43329,2,''),(270403,43329,3,''),(270404,43329,4,''),(270405,43329,5,''),(270406,43329,6,''),(270407,43329,7,''),(270408,43329,8,''),(270409,43329,9,''),(270410,43329,10,''),(270411,43329,11,''),(270412,43329,12,''),(270413,43329,13,''),(270414,43329,14,''),(270415,43329,15,''),(270416,43329,16,''),(270417,43329,17,''),(270418,43329,18,''),(270419,43329,19,''),(270420,43329,20,''),(270421,43377,1,''),(270422,43377,2,''),(270423,43377,3,''),(270424,43377,4,''),(270425,43377,5,''),(270426,43377,6,''),(270427,43377,7,''),(270428,43377,8,''),(270429,43377,9,''),(270430,43377,10,''),(270431,43377,11,''),(270432,43377,12,''),(270433,43377,13,''),(270434,43377,14,''),(270435,43377,15,''),(270436,43377,16,''),(270437,43377,17,''),(270438,43377,18,''),(270439,43377,19,''),(270440,43377,20,''),(270441,43381,1,''),(270442,43381,2,''),(270443,43381,3,''),(270444,43381,4,''),(270445,43381,5,''),(270446,43381,6,''),(270447,43381,7,''),(270448,43381,8,''),(270449,43381,9,''),(270450,43381,10,''),(270451,43381,11,''),(270452,43381,12,''),(270453,43381,13,''),(270454,43381,14,''),(270455,43381,15,''),(270456,43381,16,''),(270457,43381,17,''),(270458,43381,18,''),(270459,43381,19,''),(270460,43381,20,''),(270461,43394,1,''),(270462,43394,2,''),(270463,43394,3,''),(270464,43394,4,''),(270465,43394,5,''),(270466,43394,6,''),(270467,43394,7,''),(270468,43394,8,''),(270469,43394,9,''),(270470,43394,10,''),(270471,43394,11,''),(270472,43394,12,''),(270473,43394,13,''),(270474,43394,14,''),(270475,43394,15,''),(270476,43394,16,''),(270477,43394,17,''),(270478,43394,18,''),(270479,43394,19,''),(270480,43394,20,''),(270481,43395,1,''),(270482,43395,2,''),(270483,43395,3,''),(270484,43395,4,''),(270485,43395,5,''),(270486,43395,6,''),(270487,43395,7,''),(270488,43395,8,''),(270489,43395,9,''),(270490,43395,10,''),(270491,43395,11,''),(270492,43395,12,''),(270493,43395,13,''),(270494,43395,14,''),(270495,43395,15,''),(270496,43395,16,''),(270497,43395,17,''),(270498,43395,18,''),(270499,43395,19,''),(270500,43395,20,''),(270501,43397,1,''),(270502,43397,2,''),(270503,43397,3,''),(270504,43397,4,''),(270505,43397,5,''),(270506,43397,6,''),(270507,43397,7,''),(270508,43397,8,''),(270509,43397,9,''),(270510,43397,10,''),(270511,43397,11,''),(270512,43397,12,''),(270513,43397,13,''),(270514,43397,14,''),(270515,43397,15,''),(270516,43397,16,''),(270517,43397,17,''),(270518,43397,18,''),(270519,43397,19,''),(270520,43397,20,''),(270521,43430,1,''),(270522,43430,2,''),(270523,43430,3,''),(270524,43430,4,''),(270525,43430,5,''),(270526,43430,6,''),(270527,43430,7,''),(270528,43430,8,''),(270529,43430,9,''),(270530,43430,10,''),(270531,43430,11,''),(270532,43430,12,''),(270533,43430,13,''),(270534,43430,14,''),(270535,43430,15,''),(270536,43430,16,''),(270537,43430,17,''),(270538,43430,18,''),(270539,43430,19,''),(270540,43430,20,''),(270541,43459,1,''),(270542,43459,2,''),(270543,43459,3,''),(270544,43459,4,''),(270545,43459,5,''),(270546,43459,6,''),(270547,43459,7,''),(270548,43459,8,''),(270549,43459,9,''),(270550,43459,10,''),(270551,43459,11,''),(270552,43459,12,''),(270553,43459,13,''),(270554,43459,14,''),(270555,43459,15,''),(270556,43459,16,''),(270557,43459,17,''),(270558,43459,18,''),(270559,43459,19,''),(270560,43459,20,''),(270561,43469,1,''),(270562,43469,2,''),(270563,43469,3,''),(270564,43469,4,''),(270565,43469,5,''),(270566,43469,6,''),(270567,43469,7,''),(270568,43469,8,''),(270569,43469,9,''),(270570,43469,10,''),(270571,43469,11,''),(270572,43469,12,''),(270573,43469,13,''),(270574,43469,14,''),(270575,43469,15,''),(270576,43469,16,''),(270577,43469,17,''),(270578,43469,18,''),(270579,43469,19,''),(270580,43469,20,''),(270581,43487,1,''),(270582,43487,2,''),(270583,43487,3,''),(270584,43487,4,''),(270585,43487,5,''),(270586,43487,6,''),(270587,43487,7,''),(270588,43487,8,''),(270589,43487,9,''),(270590,43487,10,''),(270591,43487,11,''),(270592,43487,12,''),(270593,43487,13,''),(270594,43487,14,''),(270595,43487,15,''),(270596,43487,16,''),(270597,43487,17,''),(270598,43487,18,''),(270599,43487,19,''),(270600,43487,20,''),(270601,43505,1,''),(270602,43505,2,''),(270603,43505,3,''),(270604,43505,4,''),(270605,43505,5,''),(270606,43505,6,''),(270607,43505,7,''),(270608,43505,8,''),(270609,43505,9,''),(270610,43505,10,''),(270611,43505,11,''),(270612,43505,12,''),(270613,43505,13,''),(270614,43505,14,''),(270615,43505,15,''),(270616,43505,16,''),(270617,43505,17,''),(270618,43505,18,''),(270619,43505,19,''),(270620,43505,20,''),(270621,43722,1,''),(270622,43722,2,''),(270623,43722,3,''),(270624,43722,4,''),(270625,43722,5,''),(270626,43722,6,''),(270627,43722,7,''),(270628,43722,8,''),(270629,43722,9,''),(270630,43722,10,''),(270631,43722,11,''),(270632,43722,12,''),(270633,43722,13,''),(270634,43722,14,''),(270635,43722,15,''),(270636,43722,16,''),(270637,43722,17,''),(270638,43722,18,''),(270639,43722,19,''),(270640,43722,20,''),(270641,43723,1,''),(270642,43723,2,''),(270643,43723,3,''),(270644,43723,4,''),(270645,43723,5,''),(270646,43723,6,''),(270647,43723,7,''),(270648,43723,8,''),(270649,43723,9,''),(270650,43723,10,''),(270651,43723,11,''),(270652,43723,12,''),(270653,43723,13,''),(270654,43723,14,''),(270655,43723,15,''),(270656,43723,16,''),(270657,43723,17,''),(270658,43723,18,''),(270659,43723,19,''),(270660,43723,20,''),(270661,43724,1,''),(270662,43724,2,''),(270663,43724,3,''),(270664,43724,4,''),(270665,43724,5,''),(270666,43724,6,''),(270667,43724,7,''),(270668,43724,8,''),(270669,43724,9,''),(270670,43724,10,''),(270671,43724,11,''),(270672,43724,12,''),(270673,43724,13,''),(270674,43724,14,''),(270675,43724,15,''),(270676,43724,16,''),(270677,43724,17,''),(270678,43724,18,''),(270679,43724,19,''),(270680,43724,20,''),(270681,43725,1,''),(270682,43725,2,''),(270683,43725,3,''),(270684,43725,4,''),(270685,43725,5,''),(270686,43725,6,''),(270687,43725,7,''),(270688,43725,8,''),(270689,43725,9,''),(270690,43725,10,''),(270691,43725,11,''),(270692,43725,12,''),(270693,43725,13,''),(270694,43725,14,''),(270695,43725,15,''),(270696,43725,16,''),(270697,43725,17,''),(270698,43725,18,''),(270699,43725,19,''),(270700,43725,20,''),(270701,43726,1,''),(270702,43726,2,''),(270703,43726,3,''),(270704,43726,4,''),(270705,43726,5,''),(270706,43726,6,''),(270707,43726,7,''),(270708,43726,8,''),(270709,43726,9,''),(270710,43726,10,''),(270711,43726,11,''),(270712,43726,12,''),(270713,43726,13,''),(270714,43726,14,''),(270715,43726,15,''),(270716,43726,16,''),(270717,43726,17,''),(270718,43726,18,''),(270719,43726,19,''),(270720,43726,20,''),(270721,43727,1,''),(270722,43727,2,''),(270723,43727,3,''),(270724,43727,4,''),(270725,43727,5,''),(270726,43727,6,''),(270727,43727,7,''),(270728,43727,8,''),(270729,43727,9,''),(270730,43727,10,''),(270731,43727,11,''),(270732,43727,12,''),(270733,43727,13,''),(270734,43727,14,''),(270735,43727,15,''),(270736,43727,16,''),(270737,43727,17,''),(270738,43727,18,''),(270739,43727,19,''),(270740,43727,20,''),(270741,43728,1,''),(270742,43728,2,''),(270743,43728,3,''),(270744,43728,4,''),(270745,43728,5,''),(270746,43728,6,''),(270747,43728,7,''),(270748,43728,8,''),(270749,43728,9,''),(270750,43728,10,''),(270751,43728,11,''),(270752,43728,12,''),(270753,43728,13,''),(270754,43728,14,''),(270755,43728,15,''),(270756,43728,16,''),(270757,43728,17,''),(270758,43728,18,''),(270759,43728,19,''),(270760,43728,20,''),(270761,43729,1,''),(270762,43729,2,''),(270763,43729,3,''),(270764,43729,4,''),(270765,43729,5,''),(270766,43729,6,''),(270767,43729,7,''),(270768,43729,8,''),(270769,43729,9,''),(270770,43729,10,''),(270771,43729,11,''),(270772,43729,12,''),(270773,43729,13,''),(270774,43729,14,''),(270775,43729,15,''),(270776,43729,16,''),(270777,43729,17,''),(270778,43729,18,''),(270779,43729,19,''),(270780,43729,20,''),(270781,43733,1,''),(270782,43733,2,''),(270783,43733,3,''),(270784,43733,4,''),(270785,43733,5,''),(270786,43733,6,''),(270787,43733,7,''),(270788,43733,8,''),(270789,43733,9,''),(270790,43733,10,''),(270791,43733,11,''),(270792,43733,12,''),(270793,43733,13,''),(270794,43733,14,''),(270795,43733,15,''),(270796,43733,16,''),(270797,43733,17,''),(270798,43733,18,''),(270799,43733,19,''),(270800,43733,20,''),(270801,43745,1,''),(270802,43745,2,''),(270803,43745,3,''),(270804,43745,4,''),(270805,43745,5,''),(270806,43745,6,''),(270807,43745,7,''),(270808,43745,8,''),(270809,43745,9,''),(270810,43745,10,''),(270811,43745,11,''),(270812,43745,12,''),(270813,43745,13,''),(270814,43745,14,''),(270815,43745,15,''),(270816,43745,16,''),(270817,43745,17,''),(270818,43745,18,''),(270819,43745,19,''),(270820,43745,20,''),(270821,43746,1,''),(270822,43746,2,''),(270823,43746,3,''),(270824,43746,4,''),(270825,43746,5,''),(270826,43746,6,''),(270827,43746,7,''),(270828,43746,8,''),(270829,43746,9,''),(270830,43746,10,''),(270831,43746,11,''),(270832,43746,12,''),(270833,43746,13,''),(270834,43746,14,''),(270835,43746,15,''),(270836,43746,16,''),(270837,43746,17,''),(270838,43746,18,''),(270839,43746,19,''),(270840,43746,20,''),(270841,43747,1,''),(270842,43747,2,''),(270843,43747,3,''),(270844,43747,4,''),(270845,43747,5,''),(270846,43747,6,''),(270847,43747,7,''),(270848,43747,8,''),(270849,43747,9,''),(270850,43747,10,''),(270851,43747,11,''),(270852,43747,12,''),(270853,43747,13,''),(270854,43747,14,''),(270855,43747,15,''),(270856,43747,16,''),(270857,43747,17,''),(270858,43747,18,''),(270859,43747,19,''),(270860,43747,20,''),(270861,43748,1,''),(270862,43748,2,''),(270863,43748,3,''),(270864,43748,4,''),(270865,43748,5,''),(270866,43748,6,''),(270867,43748,7,''),(270868,43748,8,''),(270869,43748,9,''),(270870,43748,10,''),(270871,43748,11,''),(270872,43748,12,''),(270873,43748,13,''),(270874,43748,14,''),(270875,43748,15,''),(270876,43748,16,''),(270877,43748,17,''),(270878,43748,18,''),(270879,43748,19,''),(270880,43748,20,''),(270881,43750,1,''),(270882,43750,2,''),(270883,43750,3,''),(270884,43750,4,''),(270885,43750,5,''),(270886,43750,6,''),(270887,43750,7,''),(270888,43750,8,''),(270889,43750,9,''),(270890,43750,10,''),(270891,43750,11,''),(270892,43750,12,''),(270893,43750,13,''),(270894,43750,14,''),(270895,43750,15,''),(270896,43750,16,''),(270897,43750,17,''),(270898,43750,18,''),(270899,43750,19,''),(270900,43750,20,''),(270901,43751,1,''),(270902,43751,2,''),(270903,43751,3,''),(270904,43751,4,''),(270905,43751,5,''),(270906,43751,6,''),(270907,43751,7,''),(270908,43751,8,''),(270909,43751,9,''),(270910,43751,10,''),(270911,43751,11,''),(270912,43751,12,''),(270913,43751,13,''),(270914,43751,14,''),(270915,43751,15,''),(270916,43751,16,''),(270917,43751,17,''),(270918,43751,18,''),(270919,43751,19,''),(270920,43751,20,''),(270921,43752,1,''),(270922,43752,2,''),(270923,43752,3,''),(270924,43752,4,''),(270925,43752,5,''),(270926,43752,6,''),(270927,43752,7,''),(270928,43752,8,''),(270929,43752,9,''),(270930,43752,10,''),(270931,43752,11,''),(270932,43752,12,''),(270933,43752,13,''),(270934,43752,14,''),(270935,43752,15,''),(270936,43752,16,''),(270937,43752,17,''),(270938,43752,18,''),(270939,43752,19,''),(270940,43752,20,''),(270941,43780,1,''),(270942,43780,2,''),(270943,43780,3,''),(270944,43780,4,''),(270945,43780,5,''),(270946,43780,6,''),(270947,43780,7,''),(270948,43780,8,''),(270949,43780,9,''),(270950,43780,10,''),(270951,43780,11,''),(270952,43780,12,''),(270953,43780,13,''),(270954,43780,14,''),(270955,43780,15,''),(270956,43780,16,''),(270957,43780,17,''),(270958,43780,18,''),(270959,43780,19,''),(270960,43780,20,''),(270961,43800,1,''),(270962,43800,2,''),(270963,43800,3,''),(270964,43800,4,''),(270965,43800,5,''),(270966,43800,6,''),(270967,43800,7,''),(270968,43800,8,''),(270969,43800,9,''),(270970,43800,10,''),(270971,43800,11,''),(270972,43800,12,''),(270973,43800,13,''),(270974,43800,14,''),(270975,43800,15,''),(270976,43800,16,''),(270977,43800,17,''),(270978,43800,18,''),(270979,43800,19,''),(270980,43800,20,''),(270981,43801,1,''),(270982,43801,2,''),(270983,43801,3,''),(270984,43801,4,''),(270985,43801,5,''),(270986,43801,6,''),(270987,43801,7,''),(270988,43801,8,''),(270989,43801,9,''),(270990,43801,10,''),(270991,43801,11,''),(270992,43801,12,''),(270993,43801,13,''),(270994,43801,14,''),(270995,43801,15,''),(270996,43801,16,''),(270997,43801,17,''),(270998,43801,18,''),(270999,43801,19,''),(271000,43801,20,''),(271001,43813,1,''),(271002,43813,2,''),(271003,43813,3,''),(271004,43813,4,''),(271005,43813,5,''),(271006,43813,6,''),(271007,43813,7,''),(271008,43813,8,''),(271009,43813,9,''),(271010,43813,10,''),(271011,43813,11,''),(271012,43813,12,''),(271013,43813,13,''),(271014,43813,14,''),(271015,43813,15,''),(271016,43813,16,''),(271017,43813,17,''),(271018,43813,18,''),(271019,43813,19,''),(271020,43813,20,''),(271021,43897,1,''),(271022,43897,2,''),(271023,43897,3,''),(271024,43897,4,''),(271025,43897,5,''),(271026,43897,6,''),(271027,43897,7,''),(271028,43897,8,''),(271029,43897,9,''),(271030,43897,10,''),(271031,43897,11,''),(271032,43897,12,''),(271033,43897,13,''),(271034,43897,14,''),(271035,43897,15,''),(271036,43897,16,''),(271037,43897,17,''),(271038,43897,18,''),(271039,43897,19,''),(271040,43897,20,''),(271041,43898,1,''),(271042,43898,2,''),(271043,43898,3,''),(271044,43898,4,''),(271045,43898,5,''),(271046,43898,6,''),(271047,43898,7,''),(271048,43898,8,''),(271049,43898,9,''),(271050,43898,10,''),(271051,43898,11,''),(271052,43898,12,''),(271053,43898,13,''),(271054,43898,14,''),(271055,43898,15,''),(271056,43898,16,''),(271057,43898,17,''),(271058,43898,18,''),(271059,43898,19,''),(271060,43898,20,''),(271061,43900,1,''),(271062,43900,2,''),(271063,43900,3,''),(271064,43900,4,''),(271065,43900,5,''),(271066,43900,6,''),(271067,43900,7,''),(271068,43900,8,''),(271069,43900,9,''),(271070,43900,10,''),(271071,43900,11,''),(271072,43900,12,''),(271073,43900,13,''),(271074,43900,14,''),(271075,43900,15,''),(271076,43900,16,''),(271077,43900,17,''),(271078,43900,18,''),(271079,43900,19,''),(271080,43900,20,''),(271081,43912,1,''),(271082,43912,2,''),(271083,43912,3,''),(271084,43912,4,''),(271085,43912,5,''),(271086,43912,6,''),(271087,43912,7,''),(271088,43912,8,''),(271089,43912,9,''),(271090,43912,10,''),(271091,43912,11,''),(271092,43912,12,''),(271093,43912,13,''),(271094,43912,14,''),(271095,43912,15,''),(271096,43912,16,''),(271097,43912,17,''),(271098,43912,18,''),(271099,43912,19,''),(271100,43912,20,''),(271101,43914,1,''),(271102,43914,2,''),(271103,43914,3,''),(271104,43914,4,''),(271105,43914,5,''),(271106,43914,6,''),(271107,43914,7,''),(271108,43914,8,''),(271109,43914,9,''),(271110,43914,10,''),(271111,43914,11,''),(271112,43914,12,''),(271113,43914,13,''),(271114,43914,14,''),(271115,43914,15,''),(271116,43914,16,''),(271117,43914,17,''),(271118,43914,18,''),(271119,43914,19,''),(271120,43914,20,''),(271121,43915,1,''),(271122,43915,2,''),(271123,43915,3,''),(271124,43915,4,''),(271125,43915,5,''),(271126,43915,6,''),(271127,43915,7,''),(271128,43915,8,''),(271129,43915,9,''),(271130,43915,10,''),(271131,43915,11,''),(271132,43915,12,''),(271133,43915,13,''),(271134,43915,14,''),(271135,43915,15,''),(271136,43915,16,''),(271137,43915,17,''),(271138,43915,18,''),(271139,43915,19,''),(271140,43915,20,''),(271141,43916,1,''),(271142,43916,2,''),(271143,43916,3,''),(271144,43916,4,''),(271145,43916,5,''),(271146,43916,6,''),(271147,43916,7,''),(271148,43916,8,''),(271149,43916,9,''),(271150,43916,10,''),(271151,43916,11,''),(271152,43916,12,''),(271153,43916,13,''),(271154,43916,14,''),(271155,43916,15,''),(271156,43916,16,''),(271157,43916,17,''),(271158,43916,18,''),(271159,43916,19,''),(271160,43916,20,''),(271161,43920,1,''),(271162,43920,2,''),(271163,43920,3,''),(271164,43920,4,''),(271165,43920,5,''),(271166,43920,6,''),(271167,43920,7,''),(271168,43920,8,''),(271169,43920,9,''),(271170,43920,10,''),(271171,43920,11,''),(271172,43920,12,''),(271173,43920,13,''),(271174,43920,14,''),(271175,43920,15,''),(271176,43920,16,''),(271177,43920,17,''),(271178,43920,18,''),(271179,43920,19,''),(271180,43920,20,''),(271181,44058,1,''),(271182,44058,2,''),(271183,44058,3,''),(271184,44058,4,''),(271185,44058,5,''),(271186,44058,6,''),(271187,44058,7,''),(271188,44058,8,''),(271189,44058,9,''),(271190,44058,10,''),(271191,44058,11,''),(271192,44058,12,''),(271193,44058,13,''),(271194,44058,14,''),(271195,44058,15,''),(271196,44058,16,''),(271197,44058,17,''),(271198,44058,18,''),(271199,44058,19,''),(271200,44058,20,''),(271201,44060,1,''),(271202,44060,2,''),(271203,44060,3,''),(271204,44060,4,''),(271205,44060,5,''),(271206,44060,6,''),(271207,44060,7,''),(271208,44060,8,''),(271209,44060,9,''),(271210,44060,10,''),(271211,44060,11,''),(271212,44060,12,''),(271213,44060,13,''),(271214,44060,14,''),(271215,44060,15,''),(271216,44060,16,''),(271217,44060,17,''),(271218,44060,18,''),(271219,44060,19,''),(271220,44060,20,''),(271221,44065,1,''),(271222,44065,2,''),(271223,44065,3,''),(271224,44065,4,''),(271225,44065,5,''),(271226,44065,6,''),(271227,44065,7,''),(271228,44065,8,''),(271229,44065,9,''),(271230,44065,10,''),(271231,44065,11,''),(271232,44065,12,''),(271233,44065,13,''),(271234,44065,14,''),(271235,44065,15,''),(271236,44065,16,''),(271237,44065,17,''),(271238,44065,18,''),(271239,44065,19,''),(271240,44065,20,''),(271241,44066,1,''),(271242,44066,2,''),(271243,44066,3,''),(271244,44066,4,''),(271245,44066,5,''),(271246,44066,6,''),(271247,44066,7,''),(271248,44066,8,''),(271249,44066,9,''),(271250,44066,10,''),(271251,44066,11,''),(271252,44066,12,''),(271253,44066,13,''),(271254,44066,14,''),(271255,44066,15,''),(271256,44066,16,''),(271257,44066,17,''),(271258,44066,18,''),(271259,44066,19,''),(271260,44066,20,''),(271261,44067,1,''),(271262,44067,2,''),(271263,44067,3,''),(271264,44067,4,''),(271265,44067,5,''),(271266,44067,6,''),(271267,44067,7,''),(271268,44067,8,''),(271269,44067,9,''),(271270,44067,10,''),(271271,44067,11,''),(271272,44067,12,''),(271273,44067,13,''),(271274,44067,14,''),(271275,44067,15,''),(271276,44067,16,''),(271277,44067,17,''),(271278,44067,18,''),(271279,44067,19,''),(271280,44067,20,''),(271281,44071,1,''),(271282,44071,2,''),(271283,44071,3,''),(271284,44071,4,''),(271285,44071,5,''),(271286,44071,6,''),(271287,44071,7,''),(271288,44071,8,''),(271289,44071,9,''),(271290,44071,10,''),(271291,44071,11,''),(271292,44071,12,''),(271293,44071,13,''),(271294,44071,14,''),(271295,44071,15,''),(271296,44071,16,''),(271297,44071,17,''),(271298,44071,18,''),(271299,44071,19,''),(271300,44071,20,''),(271301,44102,1,''),(271302,44102,2,''),(271303,44102,3,''),(271304,44102,4,''),(271305,44102,5,''),(271306,44102,6,''),(271307,44102,7,''),(271308,44102,8,''),(271309,44102,9,''),(271310,44102,10,''),(271311,44102,11,''),(271312,44102,12,''),(271313,44102,13,''),(271314,44102,14,''),(271315,44102,15,''),(271316,44102,16,''),(271317,44102,17,''),(271318,44102,18,''),(271319,44102,19,''),(271320,44102,20,''),(271321,44175,1,''),(271322,44175,2,''),(271323,44175,3,''),(271324,44175,4,''),(271325,44175,5,''),(271326,44175,6,''),(271327,44175,7,''),(271328,44175,8,''),(271329,44175,9,''),(271330,44175,10,''),(271331,44175,11,''),(271332,44175,12,''),(271333,44175,13,''),(271334,44175,14,''),(271335,44175,15,''),(271336,44175,16,''),(271337,44175,17,''),(271338,44175,18,''),(271339,44175,19,''),(271340,44175,20,''),(271341,44177,1,''),(271342,44177,2,''),(271343,44177,3,''),(271344,44177,4,''),(271345,44177,5,''),(271346,44177,6,''),(271347,44177,7,''),(271348,44177,8,''),(271349,44177,9,''),(271350,44177,10,''),(271351,44177,11,''),(271352,44177,12,''),(271353,44177,13,''),(271354,44177,14,''),(271355,44177,15,''),(271356,44177,16,''),(271357,44177,17,''),(271358,44177,18,''),(271359,44177,19,''),(271360,44177,20,''),(271361,44178,1,''),(271362,44178,2,''),(271363,44178,3,''),(271364,44178,4,''),(271365,44178,5,''),(271366,44178,6,''),(271367,44178,7,''),(271368,44178,8,''),(271369,44178,9,''),(271370,44178,10,''),(271371,44178,11,''),(271372,44178,12,''),(271373,44178,13,''),(271374,44178,14,''),(271375,44178,15,''),(271376,44178,16,''),(271377,44178,17,''),(271378,44178,18,''),(271379,44178,19,''),(271380,44178,20,''),(271381,44180,1,''),(271382,44180,2,''),(271383,44180,3,''),(271384,44180,4,''),(271385,44180,5,''),(271386,44180,6,''),(271387,44180,7,''),(271388,44180,8,''),(271389,44180,9,''),(271390,44180,10,''),(271391,44180,11,''),(271392,44180,12,''),(271393,44180,13,''),(271394,44180,14,''),(271395,44180,15,''),(271396,44180,16,''),(271397,44180,17,''),(271398,44180,18,''),(271399,44180,19,''),(271400,44180,20,''),(271401,44190,1,''),(271402,44190,2,''),(271403,44190,3,''),(271404,44190,4,''),(271405,44190,5,''),(271406,44190,6,''),(271407,44190,7,''),(271408,44190,8,''),(271409,44190,9,''),(271410,44190,10,''),(271411,44190,11,''),(271412,44190,12,''),(271413,44190,13,''),(271414,44190,14,''),(271415,44190,15,''),(271416,44190,16,''),(271417,44190,17,''),(271418,44190,18,''),(271419,44190,19,''),(271420,44190,20,''),(271421,44191,1,''),(271422,44191,2,''),(271423,44191,3,''),(271424,44191,4,''),(271425,44191,5,''),(271426,44191,6,''),(271427,44191,7,''),(271428,44191,8,''),(271429,44191,9,''),(271430,44191,10,''),(271431,44191,11,''),(271432,44191,12,''),(271433,44191,13,''),(271434,44191,14,''),(271435,44191,15,''),(271436,44191,16,''),(271437,44191,17,''),(271438,44191,18,''),(271439,44191,19,''),(271440,44191,20,''),(271441,44192,1,''),(271442,44192,2,''),(271443,44192,3,''),(271444,44192,4,''),(271445,44192,5,''),(271446,44192,6,''),(271447,44192,7,''),(271448,44192,8,''),(271449,44192,9,''),(271450,44192,10,''),(271451,44192,11,''),(271452,44192,12,''),(271453,44192,13,''),(271454,44192,14,''),(271455,44192,15,''),(271456,44192,16,''),(271457,44192,17,''),(271458,44192,18,''),(271459,44192,19,''),(271460,44192,20,''),(271461,44193,1,''),(271462,44193,2,''),(271463,44193,3,''),(271464,44193,4,''),(271465,44193,5,''),(271466,44193,6,''),(271467,44193,7,''),(271468,44193,8,''),(271469,44193,9,''),(271470,44193,10,''),(271471,44193,11,''),(271472,44193,12,''),(271473,44193,13,''),(271474,44193,14,''),(271475,44193,15,''),(271476,44193,16,''),(271477,44193,17,''),(271478,44193,18,''),(271479,44193,19,''),(271480,44193,20,''),(271481,44194,1,''),(271482,44194,2,''),(271483,44194,3,''),(271484,44194,4,''),(271485,44194,5,''),(271486,44194,6,''),(271487,44194,7,''),(271488,44194,8,''),(271489,44194,9,''),(271490,44194,10,''),(271491,44194,11,''),(271492,44194,12,''),(271493,44194,13,''),(271494,44194,14,''),(271495,44194,15,''),(271496,44194,16,''),(271497,44194,17,''),(271498,44194,18,''),(271499,44194,19,''),(271500,44194,20,''),(271501,44195,1,''),(271502,44195,2,''),(271503,44195,3,''),(271504,44195,4,''),(271505,44195,5,''),(271506,44195,6,''),(271507,44195,7,''),(271508,44195,8,''),(271509,44195,9,''),(271510,44195,10,''),(271511,44195,11,''),(271512,44195,12,''),(271513,44195,13,''),(271514,44195,14,''),(271515,44195,15,''),(271516,44195,16,''),(271517,44195,17,''),(271518,44195,18,''),(271519,44195,19,''),(271520,44195,20,''),(271521,44196,1,''),(271522,44196,2,''),(271523,44196,3,''),(271524,44196,4,''),(271525,44196,5,''),(271526,44196,6,''),(271527,44196,7,''),(271528,44196,8,''),(271529,44196,9,''),(271530,44196,10,''),(271531,44196,11,''),(271532,44196,12,''),(271533,44196,13,''),(271534,44196,14,''),(271535,44196,15,''),(271536,44196,16,''),(271537,44196,17,''),(271538,44196,18,''),(271539,44196,19,''),(271540,44196,20,''),(271541,44197,1,''),(271542,44197,2,''),(271543,44197,3,''),(271544,44197,4,''),(271545,44197,5,''),(271546,44197,6,''),(271547,44197,7,''),(271548,44197,8,''),(271549,44197,9,''),(271550,44197,10,''),(271551,44197,11,''),(271552,44197,12,''),(271553,44197,13,''),(271554,44197,14,''),(271555,44197,15,''),(271556,44197,16,''),(271557,44197,17,''),(271558,44197,18,''),(271559,44197,19,''),(271560,44197,20,''),(271561,44207,1,''),(271562,44207,2,''),(271563,44207,3,''),(271564,44207,4,''),(271565,44207,5,''),(271566,44207,6,''),(271567,44207,7,''),(271568,44207,8,''),(271569,44207,9,''),(271570,44207,10,''),(271571,44207,11,''),(271572,44207,12,''),(271573,44207,13,''),(271574,44207,14,''),(271575,44207,15,''),(271576,44207,16,''),(271577,44207,17,''),(271578,44207,18,''),(271579,44207,19,''),(271580,44207,20,''),(271581,44208,1,''),(271582,44208,2,''),(271583,44208,3,''),(271584,44208,4,''),(271585,44208,5,''),(271586,44208,6,''),(271587,44208,7,''),(271588,44208,8,''),(271589,44208,9,''),(271590,44208,10,''),(271591,44208,11,''),(271592,44208,12,''),(271593,44208,13,''),(271594,44208,14,''),(271595,44208,15,''),(271596,44208,16,''),(271597,44208,17,''),(271598,44208,18,''),(271599,44208,19,''),(271600,44208,20,''),(271601,44209,1,''),(271602,44209,2,''),(271603,44209,3,''),(271604,44209,4,''),(271605,44209,5,''),(271606,44209,6,''),(271607,44209,7,''),(271608,44209,8,''),(271609,44209,9,''),(271610,44209,10,''),(271611,44209,11,''),(271612,44209,12,''),(271613,44209,13,''),(271614,44209,14,''),(271615,44209,15,''),(271616,44209,16,''),(271617,44209,17,''),(271618,44209,18,''),(271619,44209,19,''),(271620,44209,20,''),(271621,44293,1,''),(271622,44293,2,''),(271623,44293,3,''),(271624,44293,4,''),(271625,44293,5,''),(271626,44293,6,''),(271627,44293,7,''),(271628,44293,8,''),(271629,44293,9,''),(271630,44293,10,''),(271631,44293,11,''),(271632,44293,12,''),(271633,44293,13,''),(271634,44293,14,''),(271635,44293,15,''),(271636,44293,16,''),(271637,44293,17,''),(271638,44293,18,''),(271639,44293,19,''),(271640,44293,20,''),(271641,44294,1,''),(271642,44294,2,''),(271643,44294,3,''),(271644,44294,4,''),(271645,44294,5,''),(271646,44294,6,''),(271647,44294,7,''),(271648,44294,8,''),(271649,44294,9,''),(271650,44294,10,''),(271651,44294,11,''),(271652,44294,12,''),(271653,44294,13,''),(271654,44294,14,''),(271655,44294,15,''),(271656,44294,16,''),(271657,44294,17,''),(271658,44294,18,''),(271659,44294,19,''),(271660,44294,20,''),(271661,44295,1,''),(271662,44295,2,''),(271663,44295,3,''),(271664,44295,4,''),(271665,44295,5,''),(271666,44295,6,''),(271667,44295,7,''),(271668,44295,8,''),(271669,44295,9,''),(271670,44295,10,''),(271671,44295,11,''),(271672,44295,12,''),(271673,44295,13,''),(271674,44295,14,''),(271675,44295,15,''),(271676,44295,16,''),(271677,44295,17,''),(271678,44295,18,''),(271679,44295,19,''),(271680,44295,20,''),(271681,44320,1,''),(271682,44320,2,''),(271683,44320,3,''),(271684,44320,4,''),(271685,44320,5,''),(271686,44320,6,''),(271687,44320,7,''),(271688,44320,8,''),(271689,44320,9,''),(271690,44320,10,''),(271691,44320,11,''),(271692,44320,12,''),(271693,44320,13,''),(271694,44320,14,''),(271695,44320,15,''),(271696,44320,16,''),(271697,44320,17,''),(271698,44320,18,''),(271699,44320,19,''),(271700,44320,20,''),(271701,44321,1,''),(271702,44321,2,''),(271703,44321,3,''),(271704,44321,4,''),(271705,44321,5,''),(271706,44321,6,''),(271707,44321,7,''),(271708,44321,8,''),(271709,44321,9,''),(271710,44321,10,''),(271711,44321,11,''),(271712,44321,12,''),(271713,44321,13,''),(271714,44321,14,''),(271715,44321,15,''),(271716,44321,16,''),(271717,44321,17,''),(271718,44321,18,''),(271719,44321,19,''),(271720,44321,20,''),(271721,44322,1,''),(271722,44322,2,''),(271723,44322,3,''),(271724,44322,4,''),(271725,44322,5,''),(271726,44322,6,''),(271727,44322,7,''),(271728,44322,8,''),(271729,44322,9,''),(271730,44322,10,''),(271731,44322,11,''),(271732,44322,12,''),(271733,44322,13,''),(271734,44322,14,''),(271735,44322,15,''),(271736,44322,16,''),(271737,44322,17,''),(271738,44322,18,''),(271739,44322,19,''),(271740,44322,20,''),(271741,44433,1,''),(271742,44433,2,''),(271743,44433,3,''),(271744,44433,4,''),(271745,44433,5,''),(271746,44433,6,''),(271747,44433,7,''),(271748,44433,8,''),(271749,44433,9,''),(271750,44433,10,''),(271751,44433,11,''),(271752,44433,12,''),(271753,44433,13,''),(271754,44433,14,''),(271755,44433,15,''),(271756,44433,16,''),(271757,44433,17,''),(271758,44433,18,''),(271759,44433,19,''),(271760,44433,20,''),(271761,44485,1,''),(271762,44485,2,''),(271763,44485,3,''),(271764,44485,4,''),(271765,44485,5,''),(271766,44485,6,''),(271767,44485,7,''),(271768,44485,8,''),(271769,44485,9,''),(271770,44485,10,''),(271771,44485,11,''),(271772,44485,12,''),(271773,44485,13,''),(271774,44485,14,''),(271775,44485,15,''),(271776,44485,16,''),(271777,44485,17,''),(271778,44485,18,''),(271779,44485,19,''),(271780,44485,20,''),(271781,44525,1,''),(271782,44525,2,''),(271783,44525,3,''),(271784,44525,4,''),(271785,44525,5,''),(271786,44525,6,''),(271787,44525,7,''),(271788,44525,8,''),(271789,44525,9,''),(271790,44525,10,''),(271791,44525,11,''),(271792,44525,12,''),(271793,44525,13,''),(271794,44525,14,''),(271795,44525,15,''),(271796,44525,16,''),(271797,44525,17,''),(271798,44525,18,''),(271799,44525,19,''),(271800,44525,20,''),(271801,44610,1,''),(271802,44610,2,''),(271803,44610,3,''),(271804,44610,4,''),(271805,44610,5,''),(271806,44610,6,''),(271807,44610,7,''),(271808,44610,8,''),(271809,44610,9,''),(271810,44610,10,''),(271811,44610,11,''),(271812,44610,12,''),(271813,44610,13,''),(271814,44610,14,''),(271815,44610,15,''),(271816,44610,16,''),(271817,44610,17,''),(271818,44610,18,''),(271819,44610,19,''),(271820,44610,20,''),(271821,44617,1,''),(271822,44617,2,''),(271823,44617,3,''),(271824,44617,4,''),(271825,44617,5,''),(271826,44617,6,''),(271827,44617,7,''),(271828,44617,8,''),(271829,44617,9,''),(271830,44617,10,''),(271831,44617,11,''),(271832,44617,12,''),(271833,44617,13,''),(271834,44617,14,''),(271835,44617,15,''),(271836,44617,16,''),(271837,44617,17,''),(271838,44617,18,''),(271839,44617,19,''),(271840,44617,20,''),(271841,44618,1,''),(271842,44618,2,''),(271843,44618,3,''),(271844,44618,4,''),(271845,44618,5,''),(271846,44618,6,''),(271847,44618,7,''),(271848,44618,8,''),(271849,44618,9,''),(271850,44618,10,''),(271851,44618,11,''),(271852,44618,12,''),(271853,44618,13,''),(271854,44618,14,''),(271855,44618,15,''),(271856,44618,16,''),(271857,44618,17,''),(271858,44618,18,''),(271859,44618,19,''),(271860,44618,20,''),(271861,44656,1,''),(271862,44656,2,''),(271863,44656,3,''),(271864,44656,4,''),(271865,44656,5,''),(271866,44656,6,''),(271867,44656,7,''),(271868,44656,8,''),(271869,44656,9,''),(271870,44656,10,''),(271871,44656,11,''),(271872,44656,12,''),(271873,44656,13,''),(271874,44656,14,''),(271875,44656,15,''),(271876,44656,16,''),(271877,44656,17,''),(271878,44656,18,''),(271879,44656,19,''),(271880,44656,20,''),(271881,44657,1,''),(271882,44657,2,''),(271883,44657,3,''),(271884,44657,4,''),(271885,44657,5,''),(271886,44657,6,''),(271887,44657,7,''),(271888,44657,8,''),(271889,44657,9,''),(271890,44657,10,''),(271891,44657,11,''),(271892,44657,12,''),(271893,44657,13,''),(271894,44657,14,''),(271895,44657,15,''),(271896,44657,16,''),(271897,44657,17,''),(271898,44657,18,''),(271899,44657,19,''),(271900,44657,20,''),(271901,44689,1,''),(271902,44689,2,''),(271903,44689,3,''),(271904,44689,4,''),(271905,44689,5,''),(271906,44689,6,''),(271907,44689,7,''),(271908,44689,8,''),(271909,44689,9,''),(271910,44689,10,''),(271911,44689,11,''),(271912,44689,12,''),(271913,44689,13,''),(271914,44689,14,''),(271915,44689,15,''),(271916,44689,16,''),(271917,44689,17,''),(271918,44689,18,''),(271919,44689,19,''),(271920,44689,20,''),(271921,44690,1,''),(271922,44690,2,''),(271923,44690,3,''),(271924,44690,4,''),(271925,44690,5,''),(271926,44690,6,''),(271927,44690,7,''),(271928,44690,8,''),(271929,44690,9,''),(271930,44690,10,''),(271931,44690,11,''),(271932,44690,12,''),(271933,44690,13,''),(271934,44690,14,''),(271935,44690,15,''),(271936,44690,16,''),(271937,44690,17,''),(271938,44690,18,''),(271939,44690,19,''),(271940,44690,20,''),(271941,44698,1,''),(271942,44698,2,''),(271943,44698,3,''),(271944,44698,4,''),(271945,44698,5,''),(271946,44698,6,''),(271947,44698,7,''),(271948,44698,8,''),(271949,44698,9,''),(271950,44698,10,''),(271951,44698,11,''),(271952,44698,12,''),(271953,44698,13,''),(271954,44698,14,''),(271955,44698,15,''),(271956,44698,16,''),(271957,44698,17,''),(271958,44698,18,''),(271959,44698,19,''),(271960,44698,20,''),(271961,44718,1,''),(271962,44718,2,''),(271963,44718,3,''),(271964,44718,4,''),(271965,44718,5,''),(271966,44718,6,''),(271967,44718,7,''),(271968,44718,8,''),(271969,44718,9,''),(271970,44718,10,''),(271971,44718,11,''),(271972,44718,12,''),(271973,44718,13,''),(271974,44718,14,''),(271975,44718,15,''),(271976,44718,16,''),(271977,44718,17,''),(271978,44718,18,''),(271979,44718,19,''),(271980,44718,20,''),(271981,44747,1,''),(271982,44747,2,''),(271983,44747,3,''),(271984,44747,4,''),(271985,44747,5,''),(271986,44747,6,''),(271987,44747,7,''),(271988,44747,8,''),(271989,44747,9,''),(271990,44747,10,''),(271991,44747,11,''),(271992,44747,12,''),(271993,44747,13,''),(271994,44747,14,''),(271995,44747,15,''),(271996,44747,16,''),(271997,44747,17,''),(271998,44747,18,''),(271999,44747,19,''),(272000,44747,20,''),(272001,44748,1,''),(272002,44748,2,''),(272003,44748,3,''),(272004,44748,4,''),(272005,44748,5,''),(272006,44748,6,''),(272007,44748,7,''),(272008,44748,8,''),(272009,44748,9,''),(272010,44748,10,''),(272011,44748,11,''),(272012,44748,12,''),(272013,44748,13,''),(272014,44748,14,''),(272015,44748,15,''),(272016,44748,16,''),(272017,44748,17,''),(272018,44748,18,''),(272019,44748,19,''),(272020,44748,20,''),(272021,44749,1,''),(272022,44749,2,''),(272023,44749,3,''),(272024,44749,4,''),(272025,44749,5,''),(272026,44749,6,''),(272027,44749,7,''),(272028,44749,8,''),(272029,44749,9,''),(272030,44749,10,''),(272031,44749,11,''),(272032,44749,12,''),(272033,44749,13,''),(272034,44749,14,''),(272035,44749,15,''),(272036,44749,16,''),(272037,44749,17,''),(272038,44749,18,''),(272039,44749,19,''),(272040,44749,20,''),(272041,44849,1,''),(272042,44849,2,''),(272043,44849,3,''),(272044,44849,4,''),(272045,44849,5,''),(272046,44849,6,''),(272047,44849,7,''),(272048,44849,8,''),(272049,44849,9,''),(272050,44849,10,''),(272051,44849,11,''),(272052,44849,12,''),(272053,44849,13,''),(272054,44849,14,''),(272055,44849,15,''),(272056,44849,16,''),(272057,44849,17,''),(272058,44849,18,''),(272059,44849,19,''),(272060,44849,20,''),(272061,44850,1,''),(272062,44850,2,''),(272063,44850,3,''),(272064,44850,4,''),(272065,44850,5,''),(272066,44850,6,''),(272067,44850,7,''),(272068,44850,8,''),(272069,44850,9,''),(272070,44850,10,''),(272071,44850,11,''),(272072,44850,12,''),(272073,44850,13,''),(272074,44850,14,''),(272075,44850,15,''),(272076,44850,16,''),(272077,44850,17,''),(272078,44850,18,''),(272079,44850,19,''),(272080,44850,20,''),(272081,44852,1,''),(272082,44852,2,''),(272083,44852,3,''),(272084,44852,4,''),(272085,44852,5,''),(272086,44852,6,''),(272087,44852,7,''),(272088,44852,8,''),(272089,44852,9,''),(272090,44852,10,''),(272091,44852,11,''),(272092,44852,12,''),(272093,44852,13,''),(272094,44852,14,''),(272095,44852,15,''),(272096,44852,16,''),(272097,44852,17,''),(272098,44852,18,''),(272099,44852,19,''),(272100,44852,20,''),(272101,44854,1,''),(272102,44854,2,''),(272103,44854,3,''),(272104,44854,4,''),(272105,44854,5,''),(272106,44854,6,''),(272107,44854,7,''),(272108,44854,8,''),(272109,44854,9,''),(272110,44854,10,''),(272111,44854,11,''),(272112,44854,12,''),(272113,44854,13,''),(272114,44854,14,''),(272115,44854,15,''),(272116,44854,16,''),(272117,44854,17,''),(272118,44854,18,''),(272119,44854,19,''),(272120,44854,20,''),(272121,44855,1,''),(272122,44855,2,''),(272123,44855,3,''),(272124,44855,4,''),(272125,44855,5,''),(272126,44855,6,''),(272127,44855,7,''),(272128,44855,8,''),(272129,44855,9,''),(272130,44855,10,''),(272131,44855,11,''),(272132,44855,12,''),(272133,44855,13,''),(272134,44855,14,''),(272135,44855,15,''),(272136,44855,16,''),(272137,44855,17,''),(272138,44855,18,''),(272139,44855,19,''),(272140,44855,20,''),(272141,44894,1,''),(272142,44894,2,''),(272143,44894,3,''),(272144,44894,4,''),(272145,44894,5,''),(272146,44894,6,''),(272147,44894,7,''),(272148,44894,8,''),(272149,44894,9,''),(272150,44894,10,''),(272151,44894,11,''),(272152,44894,12,''),(272153,44894,13,''),(272154,44894,14,''),(272155,44894,15,''),(272156,44894,16,''),(272157,44894,17,''),(272158,44894,18,''),(272159,44894,19,''),(272160,44894,20,''),(272161,44918,1,''),(272162,44918,2,''),(272163,44918,3,''),(272164,44918,4,''),(272165,44918,5,''),(272166,44918,6,''),(272167,44918,7,''),(272168,44918,8,''),(272169,44918,9,''),(272170,44918,10,''),(272171,44918,11,''),(272172,44918,12,''),(272173,44918,13,''),(272174,44918,14,''),(272175,44918,15,''),(272176,44918,16,''),(272177,44918,17,''),(272178,44918,18,''),(272179,44918,19,''),(272180,44918,20,''),(272181,45092,1,''),(272182,45092,2,''),(272183,45092,3,''),(272184,45092,4,''),(272185,45092,5,''),(272186,45092,6,''),(272187,45092,7,''),(272188,45092,8,''),(272189,45092,9,''),(272190,45092,10,''),(272191,45092,11,''),(272192,45092,12,''),(272193,45092,13,''),(272194,45092,14,''),(272195,45092,15,''),(272196,45092,16,''),(272197,45092,17,''),(272198,45092,18,''),(272199,45092,19,''),(272200,45092,20,''),(272201,45214,1,''),(272202,45214,2,''),(272203,45214,3,''),(272204,45214,4,''),(272205,45214,5,''),(272206,45214,6,''),(272207,45214,7,''),(272208,45214,8,''),(272209,45214,9,''),(272210,45214,10,''),(272211,45214,11,''),(272212,45214,12,''),(272213,45214,13,''),(272214,45214,14,''),(272215,45214,15,''),(272216,45214,16,''),(272217,45214,17,''),(272218,45214,18,''),(272219,45214,19,''),(272220,45214,20,''),(272221,45231,1,''),(272222,45231,2,''),(272223,45231,3,''),(272224,45231,4,''),(272225,45231,5,''),(272226,45231,6,''),(272227,45231,7,''),(272228,45231,8,''),(272229,45231,9,''),(272230,45231,10,''),(272231,45231,11,''),(272232,45231,12,''),(272233,45231,13,''),(272234,45231,14,''),(272235,45231,15,''),(272236,45231,16,''),(272237,45231,17,''),(272238,45231,18,''),(272239,45231,19,''),(272240,45231,20,''),(272241,45232,1,''),(272242,45232,2,''),(272243,45232,3,''),(272244,45232,4,''),(272245,45232,5,''),(272246,45232,6,''),(272247,45232,7,''),(272248,45232,8,''),(272249,45232,9,''),(272250,45232,10,''),(272251,45232,11,''),(272252,45232,12,''),(272253,45232,13,''),(272254,45232,14,''),(272255,45232,15,''),(272256,45232,16,''),(272257,45232,17,''),(272258,45232,18,''),(272259,45232,19,''),(272260,45232,20,''),(272261,45234,1,''),(272262,45234,2,''),(272263,45234,3,''),(272264,45234,4,''),(272265,45234,5,''),(272266,45234,6,''),(272267,45234,7,''),(272268,45234,8,''),(272269,45234,9,''),(272270,45234,10,''),(272271,45234,11,''),(272272,45234,12,''),(272273,45234,13,''),(272274,45234,14,''),(272275,45234,15,''),(272276,45234,16,''),(272277,45234,17,''),(272278,45234,18,''),(272279,45234,19,''),(272280,45234,20,''),(272281,45253,1,''),(272282,45253,2,''),(272283,45253,3,''),(272284,45253,4,''),(272285,45253,5,''),(272286,45253,6,''),(272287,45253,7,''),(272288,45253,8,''),(272289,45253,9,''),(272290,45253,10,''),(272291,45253,11,''),(272292,45253,12,''),(272293,45253,13,''),(272294,45253,14,''),(272295,45253,15,''),(272296,45253,16,''),(272297,45253,17,''),(272298,45253,18,''),(272299,45253,19,''),(272300,45253,20,''),(272301,45256,1,''),(272302,45256,2,''),(272303,45256,3,''),(272304,45256,4,''),(272305,45256,5,''),(272306,45256,6,''),(272307,45256,7,''),(272308,45256,8,''),(272309,45256,9,''),(272310,45256,10,''),(272311,45256,11,''),(272312,45256,12,''),(272313,45256,13,''),(272314,45256,14,''),(272315,45256,15,''),(272316,45256,16,''),(272317,45256,17,''),(272318,45256,18,''),(272319,45256,19,''),(272320,45256,20,''),(272321,45257,1,''),(272322,45257,2,''),(272323,45257,3,''),(272324,45257,4,''),(272325,45257,5,''),(272326,45257,6,''),(272327,45257,7,''),(272328,45257,8,''),(272329,45257,9,''),(272330,45257,10,''),(272331,45257,11,''),(272332,45257,12,''),(272333,45257,13,''),(272334,45257,14,''),(272335,45257,15,''),(272336,45257,16,''),(272337,45257,17,''),(272338,45257,18,''),(272339,45257,19,''),(272340,45257,20,''),(272341,45306,1,''),(272342,45306,2,''),(272343,45306,3,''),(272344,45306,4,''),(272345,45306,5,''),(272346,45306,6,''),(272347,45306,7,''),(272348,45306,8,''),(272349,45306,9,''),(272350,45306,10,''),(272351,45306,11,''),(272352,45306,12,''),(272353,45306,13,''),(272354,45306,14,''),(272355,45306,15,''),(272356,45306,16,''),(272357,45306,17,''),(272358,45306,18,''),(272359,45306,19,''),(272360,45306,20,''),(272361,45321,1,''),(272362,45321,2,''),(272363,45321,3,''),(272364,45321,4,''),(272365,45321,5,''),(272366,45321,6,''),(272367,45321,7,''),(272368,45321,8,''),(272369,45321,9,''),(272370,45321,10,''),(272371,45321,11,''),(272372,45321,12,''),(272373,45321,13,''),(272374,45321,14,''),(272375,45321,15,''),(272376,45321,16,''),(272377,45321,17,''),(272378,45321,18,''),(272379,45321,19,''),(272380,45321,20,''),(272381,45324,1,''),(272382,45324,2,''),(272383,45324,3,''),(272384,45324,4,''),(272385,45324,5,''),(272386,45324,6,''),(272387,45324,7,''),(272388,45324,8,''),(272389,45324,9,''),(272390,45324,10,''),(272391,45324,11,''),(272392,45324,12,''),(272393,45324,13,''),(272394,45324,14,''),(272395,45324,15,''),(272396,45324,16,''),(272397,45324,17,''),(272398,45324,18,''),(272399,45324,19,''),(272400,45324,20,''),(272401,45325,1,''),(272402,45325,2,''),(272403,45325,3,''),(272404,45325,4,''),(272405,45325,5,''),(272406,45325,6,''),(272407,45325,7,''),(272408,45325,8,''),(272409,45325,9,''),(272410,45325,10,''),(272411,45325,11,''),(272412,45325,12,''),(272413,45325,13,''),(272414,45325,14,''),(272415,45325,15,''),(272416,45325,16,''),(272417,45325,17,''),(272418,45325,18,''),(272419,45325,19,''),(272420,45325,20,''),(272421,45361,1,''),(272422,45361,2,''),(272423,45361,3,''),(272424,45361,4,''),(272425,45361,5,''),(272426,45361,6,''),(272427,45361,7,''),(272428,45361,8,''),(272429,45361,9,''),(272430,45361,10,''),(272431,45361,11,''),(272432,45361,12,''),(272433,45361,13,''),(272434,45361,14,''),(272435,45361,15,''),(272436,45361,16,''),(272437,45361,17,''),(272438,45361,18,''),(272439,45361,19,''),(272440,45361,20,''),(272441,45362,1,''),(272442,45362,2,''),(272443,45362,3,''),(272444,45362,4,''),(272445,45362,5,''),(272446,45362,6,''),(272447,45362,7,''),(272448,45362,8,''),(272449,45362,9,''),(272450,45362,10,''),(272451,45362,11,''),(272452,45362,12,''),(272453,45362,13,''),(272454,45362,14,''),(272455,45362,15,''),(272456,45362,16,''),(272457,45362,17,''),(272458,45362,18,''),(272459,45362,19,''),(272460,45362,20,''),(272461,45364,1,''),(272462,45364,2,''),(272463,45364,3,''),(272464,45364,4,''),(272465,45364,5,''),(272466,45364,6,''),(272467,45364,7,''),(272468,45364,8,''),(272469,45364,9,''),(272470,45364,10,''),(272471,45364,11,''),(272472,45364,12,''),(272473,45364,13,''),(272474,45364,14,''),(272475,45364,15,''),(272476,45364,16,''),(272477,45364,17,''),(272478,45364,18,''),(272479,45364,19,''),(272480,45364,20,''),(272481,45365,1,''),(272482,45365,2,''),(272483,45365,3,''),(272484,45365,4,''),(272485,45365,5,''),(272486,45365,6,''),(272487,45365,7,''),(272488,45365,8,''),(272489,45365,9,''),(272490,45365,10,''),(272491,45365,11,''),(272492,45365,12,''),(272493,45365,13,''),(272494,45365,14,''),(272495,45365,15,''),(272496,45365,16,''),(272497,45365,17,''),(272498,45365,18,''),(272499,45365,19,''),(272500,45365,20,''),(272501,45371,1,''),(272502,45371,2,''),(272503,45371,3,''),(272504,45371,4,''),(272505,45371,5,''),(272506,45371,6,''),(272507,45371,7,''),(272508,45371,8,''),(272509,45371,9,''),(272510,45371,10,''),(272511,45371,11,''),(272512,45371,12,''),(272513,45371,13,''),(272514,45371,14,''),(272515,45371,15,''),(272516,45371,16,''),(272517,45371,17,''),(272518,45371,18,''),(272519,45371,19,''),(272520,45371,20,''),(272521,45372,1,''),(272522,45372,2,''),(272523,45372,3,''),(272524,45372,4,''),(272525,45372,5,''),(272526,45372,6,''),(272527,45372,7,''),(272528,45372,8,''),(272529,45372,9,''),(272530,45372,10,''),(272531,45372,11,''),(272532,45372,12,''),(272533,45372,13,''),(272534,45372,14,''),(272535,45372,15,''),(272536,45372,16,''),(272537,45372,17,''),(272538,45372,18,''),(272539,45372,19,''),(272540,45372,20,''),(272541,45373,1,''),(272542,45373,2,''),(272543,45373,3,''),(272544,45373,4,''),(272545,45373,5,''),(272546,45373,6,''),(272547,45373,7,''),(272548,45373,8,''),(272549,45373,9,''),(272550,45373,10,''),(272551,45373,11,''),(272552,45373,12,''),(272553,45373,13,''),(272554,45373,14,''),(272555,45373,15,''),(272556,45373,16,''),(272557,45373,17,''),(272558,45373,18,''),(272559,45373,19,''),(272560,45373,20,''),(272561,45374,1,''),(272562,45374,2,''),(272563,45374,3,''),(272564,45374,4,''),(272565,45374,5,''),(272566,45374,6,''),(272567,45374,7,''),(272568,45374,8,''),(272569,45374,9,''),(272570,45374,10,''),(272571,45374,11,''),(272572,45374,12,''),(272573,45374,13,''),(272574,45374,14,''),(272575,45374,15,''),(272576,45374,16,''),(272577,45374,17,''),(272578,45374,18,''),(272579,45374,19,''),(272580,45374,20,''),(272581,45381,1,''),(272582,45381,2,''),(272583,45381,3,''),(272584,45381,4,''),(272585,45381,5,''),(272586,45381,6,''),(272587,45381,7,''),(272588,45381,8,''),(272589,45381,9,''),(272590,45381,10,''),(272591,45381,11,''),(272592,45381,12,''),(272593,45381,13,''),(272594,45381,14,''),(272595,45381,15,''),(272596,45381,16,''),(272597,45381,17,''),(272598,45381,18,''),(272599,45381,19,''),(272600,45381,20,''),(272601,45382,1,''),(272602,45382,2,''),(272603,45382,3,''),(272604,45382,4,''),(272605,45382,5,''),(272606,45382,6,''),(272607,45382,7,''),(272608,45382,8,''),(272609,45382,9,''),(272610,45382,10,''),(272611,45382,11,''),(272612,45382,12,''),(272613,45382,13,''),(272614,45382,14,''),(272615,45382,15,''),(272616,45382,16,''),(272617,45382,17,''),(272618,45382,18,''),(272619,45382,19,''),(272620,45382,20,''),(272621,45408,1,''),(272622,45408,2,''),(272623,45408,3,''),(272624,45408,4,''),(272625,45408,5,''),(272626,45408,6,''),(272627,45408,7,''),(272628,45408,8,''),(272629,45408,9,''),(272630,45408,10,''),(272631,45408,11,''),(272632,45408,12,''),(272633,45408,13,''),(272634,45408,14,''),(272635,45408,15,''),(272636,45408,16,''),(272637,45408,17,''),(272638,45408,18,''),(272639,45408,19,''),(272640,45408,20,''),(272641,45409,1,''),(272642,45409,2,''),(272643,45409,3,''),(272644,45409,4,''),(272645,45409,5,''),(272646,45409,6,''),(272647,45409,7,''),(272648,45409,8,''),(272649,45409,9,''),(272650,45409,10,''),(272651,45409,11,''),(272652,45409,12,''),(272653,45409,13,''),(272654,45409,14,''),(272655,45409,15,''),(272656,45409,16,''),(272657,45409,17,''),(272658,45409,18,''),(272659,45409,19,''),(272660,45409,20,''),(272661,45445,1,''),(272662,45445,2,''),(272663,45445,3,''),(272664,45445,4,''),(272665,45445,5,''),(272666,45445,6,''),(272667,45445,7,''),(272668,45445,8,''),(272669,45445,9,''),(272670,45445,10,''),(272671,45445,11,''),(272672,45445,12,''),(272673,45445,13,''),(272674,45445,14,''),(272675,45445,15,''),(272676,45445,16,''),(272677,45445,17,''),(272678,45445,18,''),(272679,45445,19,''),(272680,45445,20,''),(272681,45446,1,''),(272682,45446,2,''),(272683,45446,3,''),(272684,45446,4,''),(272685,45446,5,''),(272686,45446,6,''),(272687,45446,7,''),(272688,45446,8,''),(272689,45446,9,''),(272690,45446,10,''),(272691,45446,11,''),(272692,45446,12,''),(272693,45446,13,''),(272694,45446,14,''),(272695,45446,15,''),(272696,45446,16,''),(272697,45446,17,''),(272698,45446,18,''),(272699,45446,19,''),(272700,45446,20,''),(272701,45493,1,''),(272702,45493,2,''),(272703,45493,3,''),(272704,45493,4,''),(272705,45493,5,''),(272706,45493,6,''),(272707,45493,7,''),(272708,45493,8,''),(272709,45493,9,''),(272710,45493,10,''),(272711,45493,11,''),(272712,45493,12,''),(272713,45493,13,''),(272714,45493,14,''),(272715,45493,15,''),(272716,45493,16,''),(272717,45493,17,''),(272718,45493,18,''),(272719,45493,19,''),(272720,45493,20,''),(272721,45495,1,''),(272722,45495,2,''),(272723,45495,3,''),(272724,45495,4,''),(272725,45495,5,''),(272726,45495,6,''),(272727,45495,7,''),(272728,45495,8,''),(272729,45495,9,''),(272730,45495,10,''),(272731,45495,11,''),(272732,45495,12,''),(272733,45495,13,''),(272734,45495,14,''),(272735,45495,15,''),(272736,45495,16,''),(272737,45495,17,''),(272738,45495,18,''),(272739,45495,19,''),(272740,45495,20,''),(272741,45505,1,''),(272742,45505,2,''),(272743,45505,3,''),(272744,45505,4,''),(272745,45505,5,''),(272746,45505,6,''),(272747,45505,7,''),(272748,45505,8,''),(272749,45505,9,''),(272750,45505,10,''),(272751,45505,11,''),(272752,45505,12,''),(272753,45505,13,''),(272754,45505,14,''),(272755,45505,15,''),(272756,45505,16,''),(272757,45505,17,''),(272758,45505,18,''),(272759,45505,19,''),(272760,45505,20,''),(272761,45515,1,''),(272762,45515,2,''),(272763,45515,3,''),(272764,45515,4,''),(272765,45515,5,''),(272766,45515,6,''),(272767,45515,7,''),(272768,45515,8,''),(272769,45515,9,''),(272770,45515,10,''),(272771,45515,11,''),(272772,45515,12,''),(272773,45515,13,''),(272774,45515,14,''),(272775,45515,15,''),(272776,45515,16,''),(272777,45515,17,''),(272778,45515,18,''),(272779,45515,19,''),(272780,45515,20,''),(272781,45548,1,''),(272782,45548,2,''),(272783,45548,3,''),(272784,45548,4,''),(272785,45548,5,''),(272786,45548,6,''),(272787,45548,7,''),(272788,45548,8,''),(272789,45548,9,''),(272790,45548,10,''),(272791,45548,11,''),(272792,45548,12,''),(272793,45548,13,''),(272794,45548,14,''),(272795,45548,15,''),(272796,45548,16,''),(272797,45548,17,''),(272798,45548,18,''),(272799,45548,19,''),(272800,45548,20,''),(272801,45549,1,''),(272802,45549,2,''),(272803,45549,3,''),(272804,45549,4,''),(272805,45549,5,''),(272806,45549,6,''),(272807,45549,7,''),(272808,45549,8,''),(272809,45549,9,''),(272810,45549,10,''),(272811,45549,11,''),(272812,45549,12,''),(272813,45549,13,''),(272814,45549,14,''),(272815,45549,15,''),(272816,45549,16,''),(272817,45549,17,''),(272818,45549,18,''),(272819,45549,19,''),(272820,45549,20,'');
/*!40000 ALTER TABLE `itemtextbonus` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `itemtextcommon`
--

DROP TABLE IF EXISTS `itemtextcommon`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `itemtextcommon` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `NumCard` int(11) NOT NULL,
  `NumLine` smallint(6) DEFAULT '1',
  `Text` varchar(100) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `commontext_idx` (`NumCard`,`NumLine`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemtextcommon`
--

LOCK TABLES `itemtextcommon` WRITE;
/*!40000 ALTER TABLE `itemtextcommon` DISABLE KEYS */;
/*!40000 ALTER TABLE `itemtextcommon` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `lcustomerorders`
--

DROP TABLE IF EXISTS `lcustomerorders`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `lcustomerorders` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `OrderNumber` varchar(12) NOT NULL,
  `Item` int(11) NOT NULL DEFAULT '0',
  `Section` decimal(3,0) DEFAULT '0',
  `Quantity` decimal(8,3) DEFAULT '0.000',
  `Format` varchar(10) NOT NULL DEFAULT '',
  `Price` decimal(8,2) NOT NULL DEFAULT '0.00',
  `Remark` varchar(60) NOT NULL DEFAULT '',
  `Weight` decimal(5,3) DEFAULT '0.000',
  `Units` int(11) DEFAULT '0',
  `Amount` decimal(10,2) DEFAULT '0.00',
  `Status` char(1) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `lctor_idx` (`OrderNumber`,`Item`),
  KEY `lhkey` (`OrderNumber`,`Section`),
  CONSTRAINT `lhkey` FOREIGN KEY (`OrderNumber`, `Section`) REFERENCES `hcustomerorders` (`Number`, `Section`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `lcustomerorders`
--

LOCK TABLES `lcustomerorders` WRITE;
/*!40000 ALTER TABLE `lcustomerorders` DISABLE KEYS */;
/*!40000 ALTER TABLE `lcustomerorders` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LCorderInserted` AFTER INSERT ON `LCustomerOrders`
FOR EACH ROW BEGIN

  UPDATE `HCustomerOrders` SET `HCustomerOrders`.`DatTim`=NOW(),
    `HCustomerOrders`.`AddedLines`=`HCustomerOrders`.`AddedLines`+IF(NEW.`Status`!=7,1,0),
    `HCustomerOrders`.`NetAmount`=`HCustomerOrders`.`NetAmount`+NEW.`Amount`,
    `HCustomerOrders`.`TotalWeight`=`HCustomerOrders`.`TotalWeight`+NEW.`Weight`, `HCustomerOrders`.`TotalUnits`=`HCustomerOrders`.`TotalUnits`+NEW.`Units`
    WHERE (`HCustomerOrders`.`Number`=NEW.`OrderNumber` AND `HCustomerOrders`.`Section`=NEW.`Section`) AND `HCustomerOrders`.`Status`=2;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LCorderUpdated` AFTER UPDATE ON `LCustomerOrders`
FOR EACH ROW BEGIN
  UPDATE `HCustomerOrders` SET `HCustomerOrders`.`DatTim`=NOW(),
    `HCustomerOrders`.`AddedLines`=`HCustomerOrders`.`AddedLines`+
      IF(((NEW.`Status`=5 AND OLD.`Status`=0) OR (NEW.`Status`=0 AND OLD.`Status`=5) OR (NEW.`Status`=7 AND OLD.`Status`=2) OR (NEW.`Status`=2 AND OLD.`Status`=7)), 0,
      IF((OLD.`Weight`=0 AND OLD.`Units`=0), 1, IF(NEW.`Weight`=0 AND NEW.`Units`=0, -1, 0))),
    `HCustomerOrders`.`NetAmount`=`HCustomerOrders`.`NetAmount`-OLD.`Amount`+NEW.`Amount`,
    `HCustomerOrders`.`TotalWeight`=`HCustomerOrders`.`TotalWeight`-OLD.`Weight`+NEW.`Weight`, `HCustomerOrders`.`TotalUnits`=`HCustomerOrders`.`TotalUnits`-OLD.`Units`+NEW.`Units`
    WHERE (`HCustomerOrders`.`Number`=OLD.`OrderNumber` AND `HCustomerOrders`.`Section`=OLD.`Section`) AND `HCustomerOrders`.`Status`=2;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `localcfgforms`
--

DROP TABLE IF EXISTS `localcfgforms`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localcfgforms` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Form` varchar(20) NOT NULL,
  `Widget` varchar(32) NOT NULL,
  `Value` varchar(50) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Form` (`Form`,`Widget`)
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localcfgforms`
--

LOCK TABLES `localcfgforms` WRITE;
/*!40000 ALTER TABLE `localcfgforms` DISABLE KEYS */;
INSERT INTO `localcfgforms` VALUES (1,'reports','DAILY_DATE_ALL','0'),(2,'reports','DAILY_DATE_TODAY','0'),(3,'reports','DAILY_DATE_SELECT','1'),(4,'reports','DAILY_DAILYPRINT','1'),(5,'reports','DAILY_VATPRINT','1'),(6,'reports','DAILY_PAYMENTPRINT','0'),(7,'reports','DAILY-DETAIL','2'),(8,'reports','DAILY_ORDER','1'),(9,'reports','DAILY_TRAINING','1'),(10,'reports','TICKET-GROUP','0'),(11,'reports','TICKET-DETAIL','3'),(12,'reports','TOTAL-GROUP','1'),(13,'reports','TOTAL-BROKEN','0'),(14,'reports','CASHDRAWER-GROUP','0'),(15,'reports','CASHDRAWER-BROKEN','0');
/*!40000 ALTER TABLE `localcfgforms` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localcfgscale`
--

DROP TABLE IF EXISTS `localcfgscale`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localcfgscale` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `ParamKey` varchar(128) NOT NULL DEFAULT '',
  `ParamValue` varchar(128) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `ParamKey` (`ParamKey`),
  UNIQUE KEY `ParamKey_3` (`ParamKey`)
) ENGINE=InnoDB AUTO_INCREMENT=3770 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localcfgscale`
--

LOCK TABLES `localcfgscale` WRITE;
/*!40000 ALTER TABLE `localcfgscale` DISABLE KEYS */;
INSERT INTO `localcfgscale` VALUES (8,'RESOURCES/user-icons-path','../share/pcscale/resources/icons/item'),(12,'TOUCH/loader','1'),(14,'TOUCH/scale-icon','neutral'),(15,'TOUCH/slave-icon','slave'),(796,'NETWORK/scale-number','1'),(1540,'TOUCH/vendor-icon','vnd1'),(1575,'TOUCH/items-dnd','0'),(1628,'serial-number','0'),(1631,'TOUCH/name','NeutralStyle'),(1632,'sw-full-version','2.20 - Methone. Build id: 20180919 - 3003'),(1663,'CORE/mup-group','1'),(1664,'NETWORK/scale-counter','4'),(1665,'CORE/plu-board-group','4'),(1666,'local-ticket/StartNumber','0'),(1667,'local-ticket/MaximumNumber','0'),(1681,'PRINTER/lbl-format','J45X57_FORNES'),(1682,'PRINTER/prn-format','ticket_FORNES'),(1697,'TOUCH/language','es_ES'),(1698,'TOUCH/background','waterdrops'),(2059,'PRINTER/lbl-force-continuous','0'),(2060,'PRINTER/lbl-rotate','0'),(2226,'CORE/key-groups-linked','1'),(2230,'simplified-invoice-serial',''),(3631,'TOUCH/groups-auto','0'),(3758,'TOUCH/GROUPS/menu/default','menu01'),(3759,'TOUCH/GROUPS/scroll-sales/default','kbd-PLUBoard-NeutralStyle1'),(3760,'TOUCH/GROUPS/VendorsGroup/default','Vnd-board-NeutralStyle1'),(3769,'backup-path','/home/pcscale/scale/backup');
/*!40000 ALTER TABLE `localcfgscale` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalCfgInserted` AFTER INSERT ON `LocalCfgScale`
FOR EACH ROW BEGIN
  SET @trx := (SELECT trx_new("BUTTON_UPDATE", "GUI"));
  SET @ret := (SELECT trx_set_param(@trx, "KEYNAME", CONCAT("LOCALCFGSCALE_", NEW.`ParamKey`)));
  SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", NEW.`ParamValue`));
  SET @ret := (SELECT trx_send(@trx));
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalCfgUpdated` AFTER UPDATE ON `LocalCfgScale`
FOR EACH ROW BEGIN
  IF(OLD.`ParamValue`!=NEW.`ParamValue`) THEN
    SET @trx := (SELECT trx_new("BUTTON_UPDATE", "GUI"));
    SET @ret := (SELECT trx_set_param(@trx, "KEYNAME", CONCAT("LOCALCFGSCALE_", NEW.`ParamKey`)));
    SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", NEW.`ParamValue`));
    SET @ret := (SELECT trx_send(@trx));
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalCfgDeleted` AFTER DELETE ON `LocalCfgScale`
FOR EACH ROW BEGIN
  SET @trx := (SELECT trx_new("BUTTON_UPDATE", "GUI"));
  SET @ret := (SELECT trx_set_param(@trx, "KEYNAME", CONCAT("LOCALCFGSCALE_", OLD.`ParamKey`)));
  SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", ""));
  SET @ret := (SELECT trx_send(@trx));
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `localcountries`
--

DROP TABLE IF EXISTS `localcountries`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localcountries` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` int(11) DEFAULT NULL,
  `Alpha2` char(2) NOT NULL,
  `Name` varchar(127) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Alpha2` (`Alpha2`),
  UNIQUE KEY `Code` (`Code`),
  KEY `Name` (`Name`(19))
) ENGINE=InnoDB AUTO_INCREMENT=250 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localcountries`
--

LOCK TABLES `localcountries` WRITE;
/*!40000 ALTER TABLE `localcountries` DISABLE KEYS */;
INSERT INTO `localcountries` VALUES (1,4,'AF','Afganistán'),(2,248,'AX','Islas Äland'),(3,8,'AL','Albania'),(4,12,'DZ','Algeria'),(5,16,'AS','Samoa Americana'),(6,20,'AD','Andorra'),(7,24,'AO','Angola'),(8,660,'AI','Anguila'),(9,10,'AQ','Antártida'),(10,28,'AG','Antigua y Barbuda'),(11,32,'AR','Argentina'),(12,51,'AM','Armenia'),(13,533,'AW','Aruba'),(14,36,'AU','Australia'),(15,40,'AT','Austria'),(16,31,'AZ','Azerbayán'),(17,44,'BS','Bahamas'),(18,48,'BH','Bahrein'),(19,50,'BD','Bangladesh'),(20,52,'BB','Barbados'),(21,112,'BY','Bielorrusia'),(22,56,'BE','Bélgica'),(23,84,'BZ','Belice'),(24,204,'BJ','Benín'),(25,60,'BM','Islas Bermudas'),(26,64,'BT','Bhután'),(27,68,'BO','Bolivia, Estado plurinacional de'),(28,535,'BQ','Islas BES (Caribe Neerlandés)'),(29,70,'BA','Bosnia y Herzegovina'),(30,72,'BW','Botsuana'),(31,74,'BV','Isla Bouvet'),(32,76,'BR','Brasil'),(33,86,'IO','Británico del Océano Índico, Territorio'),(34,96,'BN','Brunei Darussalam'),(35,100,'BG','Bulgaria'),(36,854,'BF','Burquina Faso'),(37,108,'BI','Burundi'),(38,116,'KH','Camboya'),(39,120,'CM','Camerún'),(40,124,'CA','Canadá'),(41,132,'CV','Cabo Verde'),(42,136,'KY','Islas Caimán'),(43,140,'CF','Centro-africana, República'),(44,148,'TD','Chad'),(45,152,'CL','Chile'),(46,156,'CN','China'),(47,162,'CX','Isla de Navidad'),(48,166,'CC','Islas Cocos (Keeling)'),(49,170,'CO','Colombia'),(50,174,'KM','Comores, Islas'),(51,178,'CG','Congo'),(52,180,'CD','Congo, República Democrática del'),(53,184,'CK','Islas Cook'),(54,188,'CR','Costa Rica'),(55,384,'CI','Costa de Marfíl'),(56,191,'HR','Croacia'),(57,192,'CU','Cuba'),(58,531,'CW','Curasao'),(59,196,'CY','Chipre'),(60,203,'CZ','República Checa'),(61,208,'DK','Dinamarca'),(62,262,'DJ','Yibuti'),(63,212,'DM','Dominica'),(64,214,'DO','República Dominicana'),(65,218,'EC','Ecuador'),(66,818,'EG','Egipto'),(67,222,'SV','El Salvador'),(68,226,'GQ','Guinea Ecuatorial'),(69,232,'ER','Eritrea'),(70,233,'EE','Estonia'),(71,231,'ET','Etiopía'),(72,238,'FK','Islas Falkland (Malvinas)'),(73,234,'FO','Islas Feroe'),(74,242,'FJ','Fiyi'),(75,246,'FI','Finlandia'),(76,250,'FR','Francia'),(77,254,'GF','Guayana Francesa'),(78,258,'PF','Polinesia Francesa'),(79,260,'TF','Territorios Franceses del Sur'),(80,266,'GA','Gabón'),(81,270,'GM','Gambia'),(82,268,'GE','Georgia'),(83,276,'DE','Alemania'),(84,288,'GH','Ghana'),(85,292,'GI','Gibraltar'),(86,300,'GR','Grecia'),(87,304,'GL','Groenlandia'),(88,308,'GD','Granada'),(89,312,'GP','Guadalupe'),(90,316,'GU','Guam'),(91,320,'GT','Guatemala'),(92,831,'GG','Guernsey'),(93,324,'GN','Guinea'),(94,624,'GW','Guinea-Bissau'),(95,328,'GY','Guyana'),(96,332,'HT','Haití'),(97,334,'HM','Islas Heard, Islas y McDonald'),(98,336,'VA','Santa Sede (Ciudad Estado del Vaticano)'),(99,340,'HN','Honduras'),(100,344,'HK','Hong Kong'),(101,348,'HU','Hungría'),(102,352,'IS','Islandia'),(103,356,'IN','India'),(104,360,'ID','Indonesia'),(105,364,'IR','Irán, República islámica de'),(106,368,'IQ','Irak'),(107,372,'IE','Irlanda'),(108,833,'IM','Isla de Man'),(109,376,'IL','Israel'),(110,380,'IT','Italia'),(111,388,'JM','Jamaica'),(112,392,'JP','Japón'),(113,832,'JE','Jersey'),(114,400,'JO','Jordania'),(115,398,'KZ','Kazajistán'),(116,404,'KE','Kenia'),(117,296,'KI','Kiribati'),(118,408,'KP','Corea, República Democrática Popular de'),(119,410,'KR','Corea, República de'),(120,414,'KW','Kuwait'),(121,417,'KG','Kirgizstán'),(122,418,'LA','República Democrática Popular de Lao'),(123,428,'LV','Letonia'),(124,422,'LB','Líbano'),(125,426,'LS','Lesoto'),(126,430,'LR','Liberia'),(127,434,'LY','Libia'),(128,438,'LI','Liechtenstein'),(129,440,'LT','Lituania'),(130,442,'LU','Luxemburgo'),(131,446,'MO','Macao'),(132,807,'MK','Macedonia, República de'),(133,450,'MG','Madagascar'),(134,454,'MW','Malawi'),(135,458,'MY','Malasia'),(136,462,'MV','Islas Maldivas'),(137,466,'ML','Mali'),(138,470,'MT','Malta'),(139,584,'MH','Islas Marshall'),(140,474,'MQ','Martinica'),(141,478,'MR','Mauritania'),(142,480,'MU','Mauricio'),(143,175,'YT','Mayotte'),(144,484,'MX','México'),(145,583,'FM','Micronesia, Estados Federados de'),(146,498,'MD','Moldavia, República de'),(147,492,'MC','Mónaco'),(148,496,'MN','Mongolia'),(149,499,'ME','Montenegro'),(150,500,'MS','Montserrat'),(151,504,'MA','Marruecos'),(152,508,'MZ','Mozambique'),(153,104,'MM','Birmania'),(154,516,'NA','Namibia'),(155,520,'NR','Nauru'),(156,524,'NP','Nepal'),(157,528,'NL','Países Bajos'),(158,540,'NC','Nueva Caledonia'),(159,554,'NZ','Nueva Zelanda'),(160,558,'NI','Nicaragua'),(161,562,'NE','Niger'),(162,566,'NG','Nigeria'),(163,570,'NU','Niue'),(164,574,'NF','Isla Norfolk'),(165,580,'MP','Islas Marianas del Norte'),(166,578,'NO','Noruega'),(167,512,'OM','Omán'),(168,586,'PK','Pakistán'),(169,585,'PW','Palau'),(170,275,'PS','Territorio palestino, Ocupado'),(171,591,'PA','Panamá'),(172,598,'PG','Papúa Nueva Guinea'),(173,600,'PY','Paraguay'),(174,604,'PE','Perú'),(175,608,'PH','Filipinas'),(176,612,'PN','Pitcairn'),(177,616,'PL','Polonia'),(178,620,'PT','Portugal'),(179,630,'PR','Puerto Rico'),(180,634,'QA','Qatar'),(181,638,'RE','Reunión'),(182,642,'RO','Rumanía'),(183,643,'RU','Federación Rusa'),(184,646,'RW','Ruanda'),(185,652,'BL','San Bartolomé'),(186,654,'SH','Santa Elena, Ascensión y Tristán de Acuña'),(187,659,'KN','San Cristobo y Nevis'),(188,662,'LC','Santa Lucía'),(189,663,'MF','San Martín (zona francesa)'),(190,666,'PM','San Pedro y Miquelon'),(191,670,'VC','San Vicente y las Granadinas'),(192,882,'WS','Samoa'),(193,674,'SM','San Marino'),(194,678,'ST','Santo Tomé y Príncipe'),(195,682,'SA','Arabia Saudí'),(196,686,'SN','Senegal'),(197,688,'RS','Serbia'),(198,690,'SC','Seychelles'),(199,694,'SL','Sierra Leona'),(200,702,'SG','Singapur'),(201,534,'SX','Isla de San Martín (zona holandsea)'),(202,703,'SK','Eslovaquia'),(203,705,'SI','Eslovenia'),(204,90,'SB','Islas Salomón'),(205,706,'SO','Somalia'),(206,710,'ZA','Suráfrica'),(207,239,'GS','Georgia del Sur e Islas Sandwitch del Sur'),(208,724,'ES','España'),(209,144,'LK','Sri Lanka'),(210,729,'SD','Sudán'),(211,740,'SR','Surinám'),(212,728,'SS','Sudán del Sur'),(213,744,'SJ','Svalbard y Jan Mayen'),(214,748,'SZ','Swazilandia'),(215,752,'SE','Suecia'),(216,756,'CH','Suiza'),(217,760,'SY','República árabe de Siria'),(218,158,'TW','Taiwán, Provincia de China'),(219,762,'TJ','Tadjikistán'),(220,834,'TZ','Tanzania, República unida de'),(221,764,'TH','Tailandia'),(222,626,'TL','Timor-Leste'),(223,768,'TG','Togo'),(224,772,'TK','Tokelau'),(225,776,'TO','Tonga'),(226,780,'TT','Trinidad y Tobago'),(227,788,'TN','Tunez'),(228,792,'TR','Turquía'),(229,795,'TM','Turkmenistán'),(230,796,'TC','Turks y Caicos, Islas'),(231,798,'TV','Tuvalu'),(232,800,'UG','Uganda'),(233,804,'UA','Ucrania'),(234,784,'AE','Emiratos Árabes Unidos'),(235,826,'GB','Reino Unido'),(236,840,'US','Estados Unidos'),(237,581,'UM','Islas menores exteriores de Estados Unidos'),(238,858,'UY','Uruguay'),(239,860,'UZ','Uzbekistán'),(240,548,'VU','Vanuatu'),(241,862,'VE','Venezuela, República Bolivariana de'),(242,704,'VN','Vietnam'),(243,92,'VG','Islas Vírgenes, Británicas'),(244,850,'VI','Islas Vírgenes, de EEUU'),(245,876,'WF','Wallis y Futuna'),(246,732,'EH','Sahara Occidental'),(247,887,'YE','Yemen'),(248,894,'ZM','Zambia'),(249,716,'ZW','Zimbabue');
/*!40000 ALTER TABLE `localcountries` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localdatalogger`
--

DROP TABLE IF EXISTS `localdatalogger`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localdatalogger` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` int(11) DEFAULT NULL,
  `DatTim` datetime DEFAULT NULL,
  `Message` varchar(255) NOT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=50 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localdatalogger`
--

LOCK TABLES `localdatalogger` WRITE;
/*!40000 ALTER TABLE `localdatalogger` DISABLE KEYS */;
INSERT INTO `localdatalogger` VALUES (22,0,'2019-02-21 15:53:06','App started. Version 2.20.3003.'),(23,0,'2019-02-21 17:50:27','App exited'),(24,0,'2019-02-21 17:51:58','App started. Version 2.20.3003.'),(25,0,'2019-02-21 18:44:57','App exited'),(26,0,'2019-02-21 18:45:43','App started. Version 2.20.3003.'),(27,0,'2019-02-22 19:24:28','App exited'),(28,0,'2019-02-25 10:40:45','App started. Version 2.20.3003.'),(29,0,'2019-02-25 11:10:03','App exited'),(30,0,'2019-02-25 17:59:15','App started. Version 2.20.3003.'),(31,0,'2019-02-25 19:07:06','App exited'),(32,0,'2019-02-26 10:39:36','App started. Version 2.20.3003.'),(33,0,'2019-02-26 19:31:33','App exited'),(34,0,'2019-02-27 09:52:33','App started. Version 2.20.3003.'),(35,0,'2019-02-27 18:01:15','App exited'),(36,0,'2019-02-27 18:05:05','App started. Version 2.20.3003.'),(37,0,'2019-02-27 18:07:04','App started. Version 2.20.3003.'),(38,0,'2019-02-27 18:08:44','App exited'),(39,0,'2019-02-27 18:13:04','App started. Version 2.20.3003.'),(40,0,'2019-02-27 18:17:16','App started. Version 2.20.3003.'),(41,0,'2019-02-27 18:18:07','App exited'),(42,0,'2019-02-27 18:18:28','App started. Version 2.20.3003.'),(43,0,'2019-02-27 18:28:16','App started. Version 2.20.3003.'),(44,0,'2019-02-27 18:38:01','App exited'),(45,0,'2019-02-28 11:16:12','App started. Version 2.20.3003.'),(46,0,'2019-02-28 17:04:52','App exited'),(47,0,'2019-02-28 17:06:29','App started. Version 2.20.3003.'),(48,0,'2019-02-28 19:15:55','App started. Version 2.20.3003.'),(49,0,'2019-02-28 19:33:59','App started. Version 2.20.3003.');
/*!40000 ALTER TABLE `localdatalogger` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localerasetickets`
--

DROP TABLE IF EXISTS `localerasetickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localerasetickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` char(1) DEFAULT 'T',
  `DatTim` datetime DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Type` (`Type`)
) ENGINE=InnoDB AUTO_INCREMENT=27 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localerasetickets`
--

LOCK TABLES `localerasetickets` WRITE;
/*!40000 ALTER TABLE `localerasetickets` DISABLE KEYS */;
INSERT INTO `localerasetickets` VALUES (24,'T','2019-02-28 19:14:00'),(25,'P','2016-10-26 10:29:17'),(26,'S','2015-11-16 08:23:51');
/*!40000 ALTER TABLE `localerasetickets` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localhtickets`
--

DROP TABLE IF EXISTS `localhtickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localhtickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `IdNetHticket` int(11) DEFAULT '0',
  `NetStat` char(1) DEFAULT '',
  `Captured` char(1) DEFAULT '',
  `CapturedDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `tckType` char(1) DEFAULT NULL,
  `Status` char(1) DEFAULT NULL,
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `TurnPrefix` char(1) DEFAULT '0',
  `Turn` int(10) unsigned DEFAULT '0',
  `TurnId` decimal(20,0) DEFAULT '0',
  `TurnDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `TurnOnTime` tinyint(1) NOT NULL DEFAULT '0',
  `SectionItem` int(11) DEFAULT '0',
  `UseSectionItem` tinyint(1) DEFAULT '0',
  `NumLines` int(11) DEFAULT '0',
  `AddedLines` int(11) DEFAULT '0',
  `GrossAmount` decimal(10,2) DEFAULT '0.00',
  `NetAmount` decimal(10,2) DEFAULT '0.00',
  `TotalWeight` decimal(12,3) DEFAULT '0.000',
  `TotalUnits` int(11) DEFAULT '0',
  `NegativeLines` int(11) DEFAULT '0',
  `NegativeAmount` decimal(10,2) DEFAULT '0.00',
  `CanceledLines` int(11) DEFAULT '0',
  `CanceledAmount` decimal(10,2) DEFAULT '0.00',
  `EAN13` char(13) NOT NULL DEFAULT '',
  `Code128` varchar(32) NOT NULL DEFAULT '',
  `StartDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `TimesPrinted` tinyint(3) unsigned NOT NULL DEFAULT '0',
  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadFrom` int(11) DEFAULT '0',
  `ReloadScale` int(11) NOT NULL DEFAULT '0',
  `ReloadDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadTicketNumber` int(11) NOT NULL DEFAULT '0',
  `ReloadtckType` char(1) DEFAULT '',
  `ReloadTo` int(11) DEFAULT '0',
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `Points` int(11) NOT NULL DEFAULT '0',
  `PointsExtra` int(11) NOT NULL DEFAULT '0',
  `PointsDiscount` int(11) NOT NULL DEFAULT '0',
  `PointsAmount` decimal(10,2) NOT NULL DEFAULT '0.00',
  `PointsChanged` int(11) NOT NULL DEFAULT '0',
  `OrderNumber` varchar(12) DEFAULT '',
  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
  `Invoice` varchar(20) NOT NULL DEFAULT '',
  `InvoiceDate` date DEFAULT '0000-00-00',
  `tckDiscount` decimal(4,2) DEFAULT '0.00',
  `ParkedNumber` int(11) DEFAULT '0',
  `NeedCopy` tinyint(1) DEFAULT '0',
  `CashStatus` char(1) CHARACTER SET utf8 COLLATE utf8_bin NOT NULL DEFAULT '-',
  `CashOperId` int(11) DEFAULT '0',
  `CashAmount` decimal(10,2) DEFAULT '0.00',
  `CashIP` int(10) unsigned NOT NULL DEFAULT '0',
  `Checksum` varchar(1024) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  KEY `Vendor` (`Vendor`),
  KEY `IdNetHticket` (`IdNetHticket`),
  KEY `Invoice` (`Invoice`),
  KEY `datim_idx` (`DatTim`),
  KEY `OrderNumber` (`OrderNumber`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localhtickets`
--

LOCK TABLES `localhtickets` WRITE;
/*!40000 ALTER TABLE `localhtickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `localhtickets` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalVendorTicketNew` AFTER INSERT ON `LocalHTickets`
FOR EACH ROW BEGIN
	REPLACE INTO `LocalVendorChanges` VALUES (NEW.`Vendor`*10, NEW.`Vendor`,'0');
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalHTicketsUpdateB` BEFORE UPDATE ON `LocalHTickets`
FOR EACH ROW BEGIN
  IF(NEW.`NetStat`='L') THEN
    IF ((OLD.`GrossAmount`!=NEW.`GrossAmount`) OR (OLD.`Customer`!=NEW.`Customer`)) THEN
      SET NEW.`PointsDiscount`=0;
      SET NEW.`PointsAmount`= 0;
    END IF;
    IF (OLD.`tckDiscount`!=NEW.`tckDiscount`) THEN
      SET NEW.`PointsDiscount`=0;
      SET NEW.`PointsAmount`= 0;
      SET NEW.`NetAmount`=CAST(NEW.`GrossAmount`*((100-NEW.`tckDiscount`)/100) AS DECIMAL(10,2));
    ELSE
      IF (OLD.`PointsDiscount`!=NEW.`PointsDiscount`) THEN
        SET @lypoints = (SELECT IFNULL(`ParamValue`,0) FROM `CfgSystem` WHERE `ParamKey`='loyalty-points-discount-currency');
        SET NEW.`PointsAmount`= IF(@lypoints,CAST(NEW.`PointsDiscount`/@lypoints AS DECIMAL(10,2)),0);
        SET NEW.`NetAmount`=CAST(NEW.`GrossAmount`*((100-NEW.`tckDiscount`)/100) AS DECIMAL(10,2))-NEW.`PointsAmount`;
      END IF;
    END IF;
    IF (OLD.`NetAmount`!=NEW.`NetAmount`) THEN
      SET @points = (SELECT IFNULL(`ParamValue`,0) FROM `CfgSystem` WHERE `ParamKey`='loyalty-currency-give-point');
      SET NEW.`Points`=IF(@points,FLOOR(NEW.`NetAmount`/@points),0);
    END IF;
    IF(OLD.`Status`!=NEW.`Status`) AND ((NEW.`Status`='C') OR (NEW.`Status`='X') OR (NEW.`Status`='D') OR (NEW.`Status`='R')) THEN
      SET NEW.`NeedCopy`=TRUE;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalHTicketsUpdate` AFTER UPDATE ON `LocalHTickets`
FOR EACH ROW BEGIN

  IF(NEW.`ReloadTo`!=0) AND (OLD.`Status`!= NEW.`Status`) THEN
    UPDATE `LTickets` SET `LTickets`.`ReloadTo`=NEW.`ReloadTo` WHERE `LTickets`.`IdHTicket`=OLD.`Id`;
  END IF;
  REPLACE INTO `LocalVendorChanges` VALUES (NEW.`Vendor`*10, NEW.`Vendor`,'0');
  IF(OLD.`Number`!= NEW.`Number`) THEN
    UPDATE `LocalLTickets` SET `LocalLTickets`.`PrintScale`=NEW.`Scale`, `LocalLTickets`.`PrintDateTime`=NEW.`DatTim`, `LocalLTickets`.`PrintTicketNumber`=NEW.`Number`,
           `LocalLTickets`.`PrinttckType`=OLD.`tckType`, `LocalLTickets`.`TicketDiscount`=OLD.`tckDiscount` WHERE `LocalLTickets`.`IdHTicket`=OLD.`Id`;
  END IF;
  IF(NEW.`NetStat`='L') THEN
    UPDATE `LocalStatus` SET `LocalTickets` = TRUE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `localints`
--

DROP TABLE IF EXISTS `localints`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localints` (
  `i` tinyint(4) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localints`
--

LOCK TABLES `localints` WRITE;
/*!40000 ALTER TABLE `localints` DISABLE KEYS */;
INSERT INTO `localints` VALUES (0),(1),(2),(3),(4),(5),(6),(7),(8),(9);
/*!40000 ALTER TABLE `localints` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localitemgroupkeyschanges`
--

DROP TABLE IF EXISTS `localitemgroupkeyschanges`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localitemgroupkeyschanges` (
  `NumKey` decimal(2,0) NOT NULL,
  PRIMARY KEY (`NumKey`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localitemgroupkeyschanges`
--

LOCK TABLES `localitemgroupkeyschanges` WRITE;
/*!40000 ALTER TABLE `localitemgroupkeyschanges` DISABLE KEYS */;
/*!40000 ALTER TABLE `localitemgroupkeyschanges` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localltickets`
--

DROP TABLE IF EXISTS `localltickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localltickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `IdHTicket` int(11) DEFAULT NULL,
  `NetStat` char(1) DEFAULT 'L',
  `lineType` smallint(6) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `SaleFormEx` char(1) DEFAULT '-',
  `Item` int(11) DEFAULT NULL,
  `Name` varchar(127) NOT NULL DEFAULT '',
  `Section` decimal(3,0) NOT NULL DEFAULT '0',
  `Family` int(11) DEFAULT '0',
  `ItemAccess` int(11) DEFAULT '0',
  `LineScale` int(11) NOT NULL DEFAULT '0',
  `LineVendor` int(11) NOT NULL DEFAULT '0',
  `LineDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `SourceKey` char(1) NOT NULL DEFAULT '0',
  `PriceChanged` tinyint(1) NOT NULL DEFAULT '0',
  `VATCode` decimal(1,0) unsigned NOT NULL DEFAULT '0',
  `VATPercent` decimal(4,2) unsigned NOT NULL DEFAULT '0.00',
  `Weight` decimal(7,3) DEFAULT NULL,
  `Tare` decimal(7,3) DEFAULT '0.000',
  `TTare` char(1) DEFAULT '',
  `lineDiscount` decimal(4,2) DEFAULT NULL,
  `Price` decimal(8,2) DEFAULT NULL,
  `Amount` decimal(10,2) DEFAULT NULL,
  `Points` int(11) NOT NULL DEFAULT '0',
  `BovineCard` smallint(3) DEFAULT NULL,
  `EarTag` varchar(30) DEFAULT NULL,
  `BovineLot` varchar(20) DEFAULT NULL,
  `FreezingDate` date DEFAULT NULL,
  `TicketDiscount` decimal(4,2) DEFAULT '0.00',
  `PrintScale` int(11) NOT NULL DEFAULT '0',
  `PrintDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `PrintTicketNumber` int(11) NOT NULL DEFAULT '0',
  `PrinttckType` char(1) DEFAULT '',
  `ReloadScale` int(11) NOT NULL DEFAULT '0',
  `ReloadDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadTicketNumber` int(11) NOT NULL DEFAULT '0',
  `ReloadtckType` char(1) DEFAULT '',
  `ReloadTo` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  KEY `IdHTicket` (`IdHTicket`),
  CONSTRAINT `localltickets_ibfk_1` FOREIGN KEY (`IdHTicket`) REFERENCES `localhtickets` (`Id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localltickets`
--

LOCK TABLES `localltickets` WRITE;
/*!40000 ALTER TABLE `localltickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `localltickets` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalLTicketsInsertB` BEFORE INSERT ON `LocalLTickets`
FOR EACH ROW BEGIN

  IF ((NEW.`NetStat`='L') AND  (NEW.`lineType`=1)) THEN
    IF ((SELECT Status FROM `LocalHTickets` WHERE `LocalHTickets`.`Id`=NEW.`IdHTicket`)='O') THEN
      SET NEW.`LineDateTime`=NOW();
      IF ((SELECT IFNULL(`ParamValue`,0) FROM `CfgSystem` WHERE `ParamKey`='loyalty-items-extra-points')!=0) THEN
        SET NEW.`Points`=IFNULL((SELECT
          IF(`ItemsExtraPoints`.`FromQuantity`,

            IF(`ItemsExtraPoints`.`Quantity`,

              IF(`ItemsExtraPoints`.`Proportional`,

                IF(NEW.`Weight`>=`ItemsExtraPoints`.`FromQuantity`,(((NEW.`Weight`-`ItemsExtraPoints`.`FromQuantity`)/`ItemsExtraPoints`.`Quantity`) +1)*`ItemsExtraPoints`.`Points`, 0),

                  IF(NEW.`Weight`>=`ItemsExtraPoints`.`FromQuantity`,(((NEW.`Weight`-`ItemsExtraPoints`.`FromQuantity`)/`ItemsExtraPoints`.`Quantity`) DIV 1 +1)*`ItemsExtraPoints`.`Points`, 0)),

              IF(NEW.`Weight`>=`ItemsExtraPoints`.`FromQuantity`,`ItemsExtraPoints`.`Points`,0)),

            IF(`ItemsExtraPoints`.`Quantity`,

              IF(`ItemsExtraPoints`.`Proportional`,

                (NEW.`Weight`/`ItemsExtraPoints`.`Quantity`)*`ItemsExtraPoints`.`Points`,

                  IF(NEW.`Weight`>=`ItemsExtraPoints`.`Quantity`,(((NEW.`Weight`-`ItemsExtraPoints`.`Quantity`)/`ItemsExtraPoints`.`Quantity`) DIV 1 +1)*`ItemsExtraPoints`.`Points`, 0)),

              `ItemsExtraPoints`.`Points`))
        FROM `ItemsExtraPoints` WHERE `ItemsExtraPoints`.`Code`=NEW.`Item`), 0);
      END IF;

    ELSE
      UPDATE `PCScaleErrorTag` SET x=1;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalLTicketsInserted` AFTER INSERT ON `LocalLTickets`
FOR EACH ROW BEGIN

  IF(NEW.`NetStat`='L') THEN
    SET @wgt = IF(NEW.`SaleForm`=1,NEW.`Weight`,0);
    SET @units = IF(NEW.`SaleForm`=0,NEW.`Weight`,0);
    CASE(NEW.`lineType`)

      WHEN 0 THEN
        UPDATE `LocalHTickets` SET `NumLines`=`NumLines`+1, `CanceledLines`=`CanceledLines`+1, `CanceledAmount`=`CanceledAmount`+NEW.`Amount`,
          `DatTim`=NOW(), `SectionItem`=NEW.`Section` WHERE `LocalHTickets`.`Id`=NEW.`IdHTicket` AND `LocalHTickets`.`Status`='O';

      WHEN 1 THEN
        UPDATE `LocalHTickets` SET `NumLines`=`NumLines`+1, `AddedLines`=`AddedLines`+1, `GrossAmount`=`GrossAmount`+NEW.`Amount`,
               `NetAmount`=CAST(`GrossAmount`*((100-`tckDiscount`)/100) AS DECIMAL(10,2)),
               `NegativeLines`=`NegativeLines`+IF(NEW.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`+IF(NEW.`Price`<0,-NEW.`Amount`,0),
               `TotalWeight`=`TotalWeight`+IF(NEW.`Price`<0,0,@wgt), `TotalUnits`=`TotalUnits`+IF(NEW.`Price`<0,0,@units),
               `LocalHTickets`.`PointsExtra`=`LocalHTickets`.`PointsExtra`+IF(NEW.`Price`>0,NEW.`Points`,0),
               `DatTim`=NOW(), `SectionItem`=NEW.`Section` WHERE `LocalHTickets`.`Id`=NEW.`IdHTicket` AND `LocalHTickets`.`Status`='O' AND `LocalHTickets`.`NetStat`='L';

      WHEN 3 THEN
        IF(NEW.`SaleForm`=0) THEN
          UPDATE `LocalHTickets` SET `TotalWeight`=`TotalWeight`+@wgt, `TotalUnits`=`TotalUnits`+@units, `PointsChanged`=`PointsChanged`+NEW.`Price`, `DatTim`=NOW()
            WHERE `LocalHTickets`.`Id`=NEW.`IdHTicket` AND `LocalHTickets`.`Status`='O' ;
        END IF;

      ELSE BEGIN END;
    END CASE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalLTicketsUpdated` AFTER UPDATE ON `LocalLTickets`
FOR EACH ROW BEGIN
  IF((NEW.`IdHTicket`!=OLD.`IdHTicket`) OR (NEW.`Amount`!=OLD.`Amount`) OR (NEW.`VATCode`!=OLD.`VATCode`) OR (NEW.`VATPercent`!=OLD.`VATPercent`)
    OR (NEW.`SaleForm`!=OLD.`SaleForm`) OR (NEW.`Weight`!=OLD.`Weight`) OR (NEW.`Price`!=OLD.`Price`) OR (NEW.`lineDiscount`!=OLD.`lineDiscount`)) THEN
    CALL `manipulate_protected_field`;
  END IF;
  IF(NEW.`lineType`!=OLD.`lineType`) THEN
    SET @wgt = IF(OLD.`SaleForm`=1,OLD.`Weight`,0);
    SET @units = IF(NEW.`SaleForm`=0,NEW.`Weight`,0);
    CASE(NEW.`lineType`)

      WHEN 0 THEN
        IF(OLD.`lineType`=1) THEN
          UPDATE `LocalHTickets` SET `AddedLines`=`AddedLines`-1, `GrossAmount`=`GrossAmount`-OLD.`Amount`,
            `NetAmount`=CAST(`GrossAmount`*((100-`tckDiscount`)/100) AS DECIMAL(10,2)),
            `NegativeLines`=`NegativeLines`-IF(OLD.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`-IF(OLD.`Price`<0,-OLD.`Amount`,0),
            `CanceledLines`=`CanceledLines`+1, `CanceledAmount` = `CanceledAmount`+OLD.`Amount`,
            `TotalWeight`=`TotalWeight`-IF(OLD.`Price`<0,0,@wgt), `TotalUnits`=`TotalUnits`-IF(NEW.`Price`<0,0,@units),
            `LocalHTickets`.`PointsExtra`=`LocalHTickets`.`PointsExtra`-IF(OLD.`Price`>0,OLD.`Points`,0),
            `DatTim`=NOW()
          WHERE `LocalHTickets`.`Id`=OLD.`IdHTicket`;
        END IF;

      WHEN 1 THEN
        IF(OLD.`lineType`=0) THEN
          UPDATE `LocalHTickets` SET `AddedLines`=`AddedLines`+1, `GrossAmount`=`GrossAmount`+OLD.`Amount`,
            `NetAmount`=CAST(`GrossAmount`*((100-`tckDiscount`)/100) AS DECIMAL(10,2)),
            `NegativeLines`=`NegativeLines`+IF(OLD.`Price`<0,1,0), `NegativeAmount`=`NegativeAmount`+IF(OLD.`Price`<0,-OLD.`Amount`,0),
            `CanceledLines`=`CanceledLines`-1, `CanceledAmount`=`CanceledAmount`-OLD.`Amount`,
            `TotalWeight`=`TotalWeight`+IF(OLD.`Price`<0,0,@wgt), `TotalUnits`=`TotalUnits`+IF(NEW.`Price`<0,0,@units),
            `LocalHTickets`.`PointsExtra`=`LocalHTickets`.`PointsExtra`+IF(OLD.`Price`>0,OLD.`Points`,0),
            `DatTim`=NOW()
          WHERE `LocalHTickets`.`Id`=OLD.`IdHTicket`;
        END IF;

      WHEN 2 THEN
        IF((NEW.`SaleForm`=0) AND (OLD.`lineType`=3)) THEN
          UPDATE `LocalHTickets` SET `TotalWeight`=`TotalWeight`-@wgt, `TotalUnits`=`TotalUnits`-@units, `PointsChanged`=`PointsChanged`-OLD.`Price`, `DatTim`=NOW() WHERE `LocalHTickets`.`Id`=OLD.`IdHTicket`;
        END IF;

      WHEN 3 THEN
        IF((NEW.`SaleForm`=0) AND (OLD.`lineType`=2)) THEN
          UPDATE `LocalHTickets` SET `TotalWeight`=`TotalWeight`+@wgt, `TotalUnits`=`TotalUnits`+@units, `PointsChanged`=`PointsChanged`+OLD.`Price`, `DatTim`=NOW() WHERE `LocalHTickets`.`Id`=OLD.`IdHTicket`;
        END IF;

      ELSE BEGIN END;
    END CASE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `localnumtickets`
--

DROP TABLE IF EXISTS `localnumtickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localnumtickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` binary(1) DEFAULT 'T',
  `Identifier` varchar(16) NOT NULL DEFAULT '',
  `Number` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Identifier` (`Identifier`,`Type`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localnumtickets`
--

LOCK TABLES `localnumtickets` WRITE;
/*!40000 ALTER TABLE `localnumtickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `localnumtickets` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localpacketlabels`
--

DROP TABLE IF EXISTS `localpacketlabels`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localpacketlabels` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Captured` char(1) DEFAULT '',
  `CapturedDateTime` datetime NOT NULL DEFAULT '1970-01-01 00:00:00',
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `lineType` smallint(6) DEFAULT NULL,
  `Item` int(11) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `Weight` decimal(7,3) DEFAULT NULL,
  `Tare` decimal(7,3) DEFAULT '0.000',
  `TTare` char(1) DEFAULT '',
  `Price` decimal(8,2) DEFAULT NULL,
  `lineDiscount` decimal(4,2) DEFAULT NULL,
  `Amount` decimal(10,2) DEFAULT NULL,
  `PacketExt` char(1) NOT NULL DEFAULT '-',
  `BovineCard` int(11) DEFAULT NULL,
  `EarTag` varchar(30) DEFAULT NULL,
  `BovineLot` varchar(20) DEFAULT NULL,
  `FreezingDate` date DEFAULT NULL,
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `DatTim` datetime DEFAULT NULL,
  `EAN13` varchar(13) NOT NULL DEFAULT '',
  `Code128` varchar(32) NOT NULL DEFAULT '',
  `Lot` varchar(20) NOT NULL DEFAULT '',
  `Ordinal` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localpacketlabels`
--

LOCK TABLES `localpacketlabels` WRITE;
/*!40000 ALTER TABLE `localpacketlabels` DISABLE KEYS */;
/*!40000 ALTER TABLE `localpacketlabels` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localpacketpresets`
--

DROP TABLE IF EXISTS `localpacketpresets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localpacketpresets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Vendor` int(11) DEFAULT NULL,
  `Level` smallint(6) DEFAULT NULL,
  `Type` smallint(6) DEFAULT NULL,
  `Value` decimal(6,3) DEFAULT NULL,
  `Number` int(11) DEFAULT '1',
  `ItemTotal` int(11) DEFAULT '0',
  `LowerLabels` int(10) unsigned NOT NULL DEFAULT '0',
  `TotalUnits` int(10) unsigned NOT NULL DEFAULT '0',
  `TotalWeight` decimal(12,3) unsigned NOT NULL DEFAULT '0.000',
  `TotalAmount` decimal(12,2) unsigned NOT NULL DEFAULT '0.00',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Vendor` (`Vendor`,`Level`)
) ENGINE=InnoDB AUTO_INCREMENT=55 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localpacketpresets`
--

LOCK TABLES `localpacketpresets` WRITE;
/*!40000 ALTER TABLE `localpacketpresets` DISABLE KEYS */;
/*!40000 ALTER TABLE `localpacketpresets` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localpaytickets`
--

DROP TABLE IF EXISTS `localpaytickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localpaytickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `IdHTicket` int(11) DEFAULT NULL,
  `PayMethod` int(11) DEFAULT NULL,
  `Delivered` decimal(10,2) DEFAULT '0.00',
  PRIMARY KEY (`Id`),
  KEY `IdHTicket` (`IdHTicket`),
  CONSTRAINT `localpaytickets_ibfk_1` FOREIGN KEY (`IdHTicket`) REFERENCES `localhtickets` (`Id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localpaytickets`
--

LOCK TABLES `localpaytickets` WRITE;
/*!40000 ALTER TABLE `localpaytickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `localpaytickets` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localplukeyschanges`
--

DROP TABLE IF EXISTS `localplukeyschanges`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localplukeyschanges` (
  `Section` int(11) DEFAULT NULL,
  `Board` int(11) DEFAULT NULL,
  `PLUKey` int(11) DEFAULT NULL,
  `Code` int(11) DEFAULT NULL,
  UNIQUE KEY `Section` (`Section`,`Board`,`PLUKey`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localplukeyschanges`
--

LOCK TABLES `localplukeyschanges` WRITE;
/*!40000 ALTER TABLE `localplukeyschanges` DISABLE KEYS */;
/*!40000 ALTER TABLE `localplukeyschanges` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localsservicelabels`
--

DROP TABLE IF EXISTS `localsservicelabels`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localsservicelabels` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Captured` char(1) DEFAULT '',
  `CapturedDateTime` datetime NOT NULL DEFAULT '1970-01-01 00:00:00',
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `lineType` smallint(6) DEFAULT NULL,
  `Item` int(11) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `Weight` decimal(7,3) DEFAULT NULL,
  `Tare` decimal(7,3) DEFAULT '0.000',
  `TTare` char(1) DEFAULT '',
  `Price` decimal(8,2) DEFAULT NULL,
  `lineDiscount` decimal(4,2) DEFAULT NULL,
  `Amount` decimal(10,2) DEFAULT NULL,
  `PacketExt` char(1) NOT NULL DEFAULT '-',
  `BovineCard` int(11) DEFAULT NULL,
  `EarTag` varchar(30) DEFAULT NULL,
  `BovineLot` varchar(20) DEFAULT NULL,
  `FreezingDate` date DEFAULT NULL,
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `DatTim` datetime DEFAULT NULL,
  `EAN13` varchar(13) NOT NULL DEFAULT '',
  `Code128` varchar(32) NOT NULL DEFAULT '',
  `Lot` varchar(20) NOT NULL DEFAULT '',
  `Ordinal` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localsservicelabels`
--

LOCK TABLES `localsservicelabels` WRITE;
/*!40000 ALTER TABLE `localsservicelabels` DISABLE KEYS */;
/*!40000 ALTER TABLE `localsservicelabels` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localstatus`
--

DROP TABLE IF EXISTS `localstatus`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localstatus` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Version` smallint(6) DEFAULT '14',
  `AppVersion` varchar(16) DEFAULT '',
  `Language` varchar(31) NOT NULL DEFAULT '',
  `EraseTickets` tinyint(1) DEFAULT '0',
  `RemoveScale` tinyint(1) DEFAULT '0',
  `CheckCounters` tinyint(1) DEFAULT '0',
  `CheckCountersKeys` tinyint(1) DEFAULT '0',
  `CheckSections` tinyint(1) DEFAULT '0',
  `CheckFamilies` tinyint(1) DEFAULT '0',
  `CheckCounterFamilies` tinyint(1) DEFAULT '0',
  `CheckItemGroupKeysText` tinyint(1) DEFAULT '0',
  `RefreshTouchsel` tinyint(1) DEFAULT '0',
  `ResizeTouch` tinyint(1) DEFAULT '0',
  `CheckTareCodes` tinyint(1) DEFAULT '0',
  `CheckScripts` tinyint(1) DEFAULT '0',
  `LocalTickets` tinyint(1) DEFAULT '0',
  `AdvertisingChanged` tinyint(1) DEFAULT '0',
  `CheckCustomerOrders` tinyint(1) DEFAULT '0',
  `CheckCash` tinyint(1) DEFAULT '0',
  `CheckTurn` tinyint(1) DEFAULT '0',
  `CheckTotalDaily` tinyint(1) DEFAULT '1',
  `CheckTotalScaleFamilies` tinyint(1) DEFAULT '1',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localstatus`
--

LOCK TABLES `localstatus` WRITE;
/*!40000 ALTER TABLE `localstatus` DISABLE KEYS */;
INSERT INTO `localstatus` VALUES (1,111,'2.20.3003','',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1);
/*!40000 ALTER TABLE `localstatus` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localtreeboard`
--

DROP TABLE IF EXISTS `localtreeboard`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localtreeboard` (
  `Board` int(11) NOT NULL,
  `Level1` int(11) NOT NULL DEFAULT '0',
  `Level2` int(11) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Board`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localtreeboard`
--

LOCK TABLES `localtreeboard` WRITE;
/*!40000 ALTER TABLE `localtreeboard` DISABLE KEYS */;
/*!40000 ALTER TABLE `localtreeboard` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localvendorchanges`
--

DROP TABLE IF EXISTS `localvendorchanges`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localvendorchanges` (
  `Id` int(11) NOT NULL,
  `Vendor` int(11) DEFAULT NULL,
  `tck` char(1) DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localvendorchanges`
--

LOCK TABLES `localvendorchanges` WRITE;
/*!40000 ALTER TABLE `localvendorchanges` DISABLE KEYS */;
/*!40000 ALTER TABLE `localvendorchanges` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localvendorkeys`
--

DROP TABLE IF EXISTS `localvendorkeys`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localvendorkeys` (
  `Counter` int(11) DEFAULT NULL,
  `VendorKey` int(11) DEFAULT NULL,
  `Vendor` int(11) NOT NULL,
  PRIMARY KEY (`Vendor`),
  UNIQUE KEY `Counter` (`Counter`,`VendorKey`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localvendorkeys`
--

LOCK TABLES `localvendorkeys` WRITE;
/*!40000 ALTER TABLE `localvendorkeys` DISABLE KEYS */;
INSERT INTO `localvendorkeys` VALUES (2,1,52001),(2,2,52002),(2,3,52003),(2,4,52004),(2,5,52005),(2,6,52006),(2,7,52007),(2,8,52008),(2,9,52009),(2,10,52000),(3,1,50001),(3,2,50002),(3,3,50003),(3,4,50004),(3,5,50005),(3,6,50006),(3,7,50007),(3,8,50008),(3,9,50009),(3,10,50000),(4,1,54001),(4,2,54002),(4,3,54003),(4,4,54004),(4,5,54005),(4,6,54006),(4,7,54007),(4,8,54008),(4,9,54009),(4,10,54000),(5,2,58002),(5,3,58003),(5,4,58004),(5,5,58005),(5,6,58006),(5,7,58007),(5,8,58008),(5,9,58009),(5,10,58000),(6,1,56001),(6,2,56002),(6,3,56003),(6,4,56004),(6,5,56005),(6,6,56006),(6,7,56007),(6,8,56008),(6,9,56009),(6,10,56000),(8,1,59001),(8,2,59002),(8,3,59003),(8,4,59004),(8,5,59005),(8,6,59006),(8,7,59007),(8,8,59008),(8,9,59009);
/*!40000 ALTER TABLE `localvendorkeys` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertLocalVendorKey` AFTER INSERT ON `LocalVendorKeys`
FOR EACH ROW BEGIN
  IF (NEW.`Counter`= GLOBAL_GET("current_counter")) THEN
    REPLACE INTO `LocalVendorKeysChanges` (`VendorKey`, `Vendor`, `VendorRelease`,`Name`, `Icon`) 
     SELECT NEW.`VendorKey`, NEW.`Vendor`, 8880, `Vendors`.`Name`, `Vendors`.`Icon` 
     FROM `Vendors` WHERE `Vendors`.`Code`=NEW.`Vendor`;
    IF @disable_localregister IS NULL THEN
      INSERT INTO `LocalVendorRegister` (`Vendor`,`RegisterDateTime`) VALUES (NEW.`Vendor`,NOW());
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateLocalVendorKey` AFTER UPDATE ON `LocalVendorKeys`
FOR EACH ROW BEGIN
  SET @counter := GLOBAL_GET("current_counter");
  IF (OLD.`Counter`= @counter) THEN
    INSERT INTO `LocalVendorKeysChanges` (`VendorKey`, `VendorRelease`) VALUE (OLD.`VendorKey`, OLD.`Vendor` )
    ON DUPLICATE KEY UPDATE `LocalVendorKeysChanges`.`VendorRelease`=OLD.`Vendor`;
  END IF;
  IF (NEW.`Counter`= @counter) THEN
    INSERT INTO `LocalVendorKeysChanges` (`VendorKey`, `Vendor`, `VendorRelease`, `Name`, `Icon`) 
      SELECT NEW.`VendorKey`, NEW.`Vendor`, OLD.`Vendor`, `Vendors`.`Name`, `Vendors`.`Icon` FROM `Vendors` WHERE `Vendors`.`Code`=NEW.`Vendor` 
    ON DUPLICATE KEY UPDATE `LocalVendorKeysChanges`.`Vendor`=NEW.`Vendor`, `LocalVendorKeysChanges`.`Name`=VALUES(`Name`), `LocalVendorKeysChanges`.`Icon`=VALUES(`Icon`);
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteLocalVendorKey` AFTER DELETE ON `LocalVendorKeys`
FOR EACH ROW BEGIN
  IF (OLD.`Counter`= GLOBAL_GET("current_counter")) THEN
    REPLACE INTO `LocalVendorKeysChanges` (`VendorKey`, `Vendor`, `VendorRelease`) VALUE (OLD.`VendorKey`, 0, OLD.`Vendor` );
    IF @disable_localregister IS NULL THEN
      UPDATE `LocalVendorRegister` SET `ReleaseDateTime`=NOW() WHERE `Vendor`=OLD.`Vendor` AND `ReleaseDateTime`<='1970-01-01 00:00:00';
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `localvendorkeyschanges`
--

DROP TABLE IF EXISTS `localvendorkeyschanges`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localvendorkeyschanges` (
  `VendorKey` int(11) NOT NULL,
  `Vendor` int(11) DEFAULT '0',
  `VendorRelease` int(11) DEFAULT '0',
  `Name` varchar(64) DEFAULT NULL,
  `Icon` varchar(127) NOT NULL DEFAULT '',
  PRIMARY KEY (`VendorKey`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localvendorkeyschanges`
--

LOCK TABLES `localvendorkeyschanges` WRITE;
/*!40000 ALTER TABLE `localvendorkeyschanges` DISABLE KEYS */;
/*!40000 ALTER TABLE `localvendorkeyschanges` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localvendorpacket`
--

DROP TABLE IF EXISTS `localvendorpacket`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localvendorpacket` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Vendor` int(11) DEFAULT NULL,
  `Auto` tinyint(1) DEFAULT NULL,
  `ClearStart` tinyint(1) DEFAULT NULL,
  `Repetitions` smallint(6) DEFAULT '0',
  `Copies` smallint(6) DEFAULT '0',
  `EAN13` char(1) DEFAULT '1',
  `PacketExt` char(1) NOT NULL DEFAULT '-',
  `BovineCard` int(11) DEFAULT NULL,
  `EarTag` varchar(30) DEFAULT NULL,
  `BovineLot` varchar(20) DEFAULT NULL,
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `PresetON` tinyint(1) DEFAULT '0',
  `Item` int(11) DEFAULT '0',
  `SaleForm` smallint(6) DEFAULT NULL,
  `Price` decimal(8,2) DEFAULT NULL,
  `Units` int(10) unsigned NOT NULL DEFAULT '0',
  `Tare` decimal(7,3) DEFAULT NULL,
  `Discount` decimal(4,2) DEFAULT NULL,
  `Lot` varchar(20) NOT NULL DEFAULT '',
  `Ordinal` int(10) unsigned DEFAULT NULL,
  `ExpiringDate` date NOT NULL DEFAULT '0000-00-00',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Vendor` (`Vendor`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localvendorpacket`
--

LOCK TABLES `localvendorpacket` WRITE;
/*!40000 ALTER TABLE `localvendorpacket` DISABLE KEYS */;
INSERT INTO `localvendorpacket` VALUES (1,101,1,1,1,0,'1','-',0,'','','2',0,1,1,55.00,1,0.000,0.00,'',NULL,'0000-00-00'),(2,1,1,1,1,0,'1','-',0,'','','',0,55522,1,4.54,1,0.000,0.00,'',NULL,'0000-00-00'),(3,106,1,1,0,0,'1','-',0,'','','',0,1,1,55.00,1,0.095,0.00,'',0,'0000-00-00'),(4,102,1,1,0,0,'1','-',0,'','','',0,1,1,55.00,1,0.000,0.00,'',0,'0000-00-00'),(5,201,1,1,0,0,'1','-',0,'','','',0,1048,1,12.99,1,0.000,0.00,'',0,'0000-00-00'),(6,202,1,1,0,0,'1','-',0,'','','',0,1005,1,12.99,1,0.000,0.00,'',0,'0000-00-00'),(7,54001,1,1,1,0,'1','-',0,'','','',0,37800,1,6.49,1,0.000,0.00,'',0,'0000-00-00');
/*!40000 ALTER TABLE `localvendorpacket` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localvendorregister`
--

DROP TABLE IF EXISTS `localvendorregister`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localvendorregister` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Vendor` int(11) DEFAULT NULL,
  `RegisterDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `RegisterSend` smallint(6) DEFAULT '0',
  `ReleaseDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReleaseSend` smallint(6) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localvendorregister`
--

LOCK TABLES `localvendorregister` WRITE;
/*!40000 ALTER TABLE `localvendorregister` DISABLE KEYS */;
/*!40000 ALTER TABLE `localvendorregister` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `localworkingmodes`
--

DROP TABLE IF EXISTS `localworkingmodes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `localworkingmodes` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `WorkingMode` varchar(20) NOT NULL DEFAULT '',
  `ParamName` varchar(32) NOT NULL,
  `ParamValue` varchar(200) NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `WorkingMode` (`WorkingMode`,`ParamName`)
) ENGINE=InnoDB AUTO_INCREMENT=2407 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localworkingmodes`
--

LOCK TABLES `localworkingmodes` WRITE;
/*!40000 ALTER TABLE `localworkingmodes` DISABLE KEYS */;
INSERT INTO `localworkingmodes` VALUES (1,'','config/CORE/USER-CHECK-MENU-FORM','ON'),(6,'00 sales','aftersale','0'),(13,'00 sales','based','sales'),(34,'00 sales','time-aftersale','1'),(46,'01 packet','based','packet'),(60,'01 packet','return-weight','ON'),(67,'02 selfservice','amount-zero','OFF'),(68,'02 selfservice','autocode','OFF'),(69,'02 selfservice','based','selfservice'),(70,'02 selfservice','change-price','OFF'),(71,'02 selfservice','code-digits','6'),(72,'02 selfservice','constant-key','ON'),(75,'02 selfservice','item-access','1'),(77,'02 selfservice','negative-amount','OFF'),(78,'02 selfservice','pluprc-key','OFF'),(79,'02 selfservice','return-weight','ON'),(80,'02 selfservice','sections',''),(81,'02 selfservice','show-print-ss','ON'),(82,'02 selfservice','time-error-ss','3'),(84,'02 selfservice','user-check-mode','OFF'),(85,'02 selfservice','user-check-mode-form','ON'),(93,'03 pos','based','sales'),(310,'00 sales','allow-reload','ON'),(483,'01 packet','user-check-mode-form','OFF'),(484,'01 packet','user-check-mode','OFF'),(485,'01 packet','user-prog-prices','ON'),(486,'01 packet','sections',''),(487,'01 packet','packet-customer','ON'),(488,'01 packet','show-print-packet','ON'),(489,'01 packet','allow-mult-key','ON'),(490,'01 packet','packet-vendor','0'),(491,'01 packet','show-form-packet','0'),(492,'01 packet','packet-after-item','OFF'),(493,'01 packet','autocode','ON'),(494,'01 packet','pluprc-key','ON'),(495,'01 packet','constant-key','ON'),(496,'01 packet','change-price','ON'),(497,'01 packet','amount-zero','ON'),(498,'01 packet','negative-amount','ON'),(499,'01 packet','code-digits','6'),(500,'01 packet','item-access','0'),(501,'01 packet','mixed-allowed','OFF'),(502,'01 packet','weightlow-tsec','0'),(503,'01 packet','weightlow-beep','ON'),(504,'01 packet','weightlow-zero','OFF'),(619,'03 pos','allow-reload','ON'),(1445,'03 pos','user-check-mode-form','OFF'),(1446,'03 pos','user-check-mode','OFF'),(1447,'03 pos','user-prog-prices','ON'),(1448,'03 pos','sections',''),(1449,'03 pos','print-ticket','ON'),(1450,'03 pos','print-label-total','OFF'),(1451,'03 pos','print-label-line','OFF'),(1452,'03 pos','close-with-line','OFF'),(1453,'03 pos','total-discount','ON'),(1454,'03 pos','sales-customer','ON'),(1455,'03 pos','allow-orders','ON'),(1456,'03 pos','payment','2'),(1457,'03 pos','startlocal','0'),(1458,'03 pos','allow-cancel','ON'),(1459,'03 pos','allow-amount0','ON'),(1460,'03 pos','auto-fix','ON'),(1461,'03 pos','times-print','250'),(1463,'03 pos','auto-vendor','ON'),(1464,'03 pos','vendor-hide-code','OFF'),(1465,'03 pos','vendor-summary','OFF'),(1466,'03 pos','customer-care','OFF'),(1467,'03 pos','park-ticket','ON'),(1468,'03 pos','ticket-format',''),(1469,'03 pos','label-format',''),(1470,'03 pos','parking-format','parked'),(1471,'03 pos','reload-type','48'),(1472,'03 pos','reprint-type','240'),(1473,'03 pos','allow-training','OFF'),(1474,'03 pos','autocode','OFF'),(1475,'03 pos','pluprc-key','ON'),(1476,'03 pos','constant-key','ON'),(1477,'03 pos','change-price','ON'),(1478,'03 pos','return-weight','ON'),(1479,'03 pos','amount-zero','ON'),(1480,'03 pos','negative-amount','ON'),(1481,'03 pos','code-digits','6'),(1482,'03 pos','item-access','0'),(1483,'03 pos','pos-method','ON'),(1484,'03 pos','clear-tare','ON'),(1485,'03 pos','weightlow-tsec','0'),(1486,'03 pos','weightlow-beep','ON'),(1487,'03 pos','weightlow-zero','OFF'),(1571,'03 pos','description','POS'),(1572,'03 pos','keyboard','POS'),(1573,'03 pos','user','admin'),(1574,'03 pos','key-fnc-on-pressed','0'),(1575,'03 pos','keyboard-loop',''),(1576,'03 pos','keyboard-vendors','POSVendor'),(1577,'03 pos','PLUload-func','2'),(1578,'03 pos','PLUload-mode','1'),(1579,'03 pos','in-use','1'),(1770,'02 selfservice','description','Autoservicio'),(1771,'02 selfservice','keyboard','selfservice'),(1772,'02 selfservice','user','admin'),(1773,'02 selfservice','key-fnc-on-pressed','1'),(1774,'02 selfservice','keyboard-loop',''),(1775,'02 selfservice','keyboard-vendors',''),(1776,'02 selfservice','PLUload-func','2'),(1777,'02 selfservice','PLUload-mode','0'),(1778,'02 selfservice','in-use','1'),(1781,'01 packet','description','Envasado'),(1782,'01 packet','keyboard','packet'),(1783,'01 packet','user','admin'),(1784,'01 packet','key-fnc-on-pressed','0'),(1785,'01 packet','keyboard-loop','packet-loop'),(1786,'01 packet','keyboard-vendors',''),(1787,'01 packet','PLUload-func','2'),(1788,'01 packet','PLUload-mode','1'),(1789,'01 packet','in-use','1'),(1938,'','working-mode','00 sales'),(1957,'','code-digits','6'),(2023,'00 sales','description','Venta'),(2024,'00 sales','keyboard','sales_TurnFull'),(2025,'00 sales','user','admin'),(2026,'00 sales','key-fnc-on-pressed','0'),(2027,'00 sales','keyboard-loop',''),(2028,'00 sales','keyboard-vendors',''),(2029,'00 sales','PLUload-func','2'),(2030,'00 sales','PLUload-mode','1'),(2031,'00 sales','in-use','1'),(2032,'00 sales','user-check-mode-form','OFF'),(2033,'00 sales','user-check-mode','OFF'),(2034,'00 sales','user-prog-prices','ON'),(2035,'00 sales','sections',''),(2036,'00 sales','print-ticket','ON'),(2037,'00 sales','print-label-total','ON'),(2038,'00 sales','print-label-line','OFF'),(2039,'00 sales','close-with-line','OFF'),(2040,'00 sales','total-discount','ON'),(2041,'00 sales','sales-customer','ON'),(2042,'00 sales','allow-orders','ON'),(2043,'00 sales','payment','0'),(2044,'00 sales','startlocal','0'),(2045,'00 sales','allow-cancel','ON'),(2046,'00 sales','allow-amount0','ON'),(2047,'00 sales','auto-fix','OFF'),(2048,'00 sales','times-print','250'),(2050,'00 sales','auto-vendor','ON'),(2051,'00 sales','vendor-hide-code','OFF'),(2052,'00 sales','vendor-summary','OFF'),(2053,'00 sales','customer-care','OFF'),(2054,'00 sales','park-ticket','ON'),(2055,'00 sales','single-section','ON'),(2056,'00 sales','ean13-section-item','ON'),(2057,'00 sales','ticket-format',''),(2058,'00 sales','label-format',''),(2059,'00 sales','parking-format','parked'),(2060,'00 sales','reload-type','48'),(2061,'00 sales','reprint-type','240'),(2062,'00 sales','allow-training','OFF'),(2063,'00 sales','hold-plusec','ON'),(2064,'00 sales','autocode','OFF'),(2065,'00 sales','pluprc-key','ON'),(2066,'00 sales','constant-key','ON'),(2067,'00 sales','change-price','ON'),(2068,'00 sales','return-weight','ON'),(2069,'00 sales','amount-zero','ON'),(2070,'00 sales','negative-amount','ON'),(2071,'00 sales','code-digits','6'),(2072,'00 sales','item-access','0'),(2073,'00 sales','pos-method','OFF'),(2074,'00 sales','clear-tare','ON'),(2075,'00 sales','weightlow-tsec','0'),(2076,'00 sales','weightlow-beep','ON'),(2077,'00 sales','weightlow-zero','OFF'),(2078,'','user-check-menu-form','OFF'),(2106,'00 sales','ticket-days','1'),(2107,'00 sales','simplified-invoice','OFF'),(2108,'00 sales','plukey-dnd','Full'),(2155,'01 packet','edit-bovine-card','OFF'),(2156,'01 packet','label-format',''),(2157,'01 packet','hold-plusec','ON'),(2158,'01 packet','plukey-dnd','Full'),(2181,'02 selfservice','label-format',''),(2182,'02 selfservice','plukey-dnd','Not allowed'),(2188,'02 selfservice','weightlow-tsec','0'),(2189,'02 selfservice','weightlow-beep','ON'),(2190,'02 selfservice','weightlow-zero','OFF'),(2198,'03 pos','single-section','OFF'),(2199,'03 pos','ean13-section-item','OFF'),(2200,'03 pos','ticket-days','1'),(2201,'03 pos','simplified-invoice','ON'),(2202,'03 pos','hold-plusec','ON'),(2203,'03 pos','plukey-dnd','Full'),(2204,'00 sales','current-data-mode','0'),(2206,'03 pos','reprint-ticket','ON'),(2207,'03 pos','reprint-label-total','OFF'),(2208,'00 sales','reprint-ticket','ON'),(2209,'00 sales','reprint-label-total','OFF'),(2210,'00 sales','hold-plu-to-set','Full'),(2211,'01 packet','hold-plu-to-set','Full'),(2212,'03 pos','hold-plu-to-set','Full'),(2213,'02 selfservice','current-data-mode','0'),(2214,'01 packet','current-data-mode','0'),(2215,'03 pos','print-ticket-suborder','ON'),(2216,'00 sales','print-ticket-suborder','ON'),(2218,'03 pos','print-label-total-suborder','OFF'),(2219,'00 sales','print-label-total-suborder','OFF'),(2221,'03 pos','print-ticket-order','ON'),(2222,'00 sales','print-ticket-order','ON'),(2224,'03 pos','print-label-order','OFF'),(2225,'00 sales','print-label-order','OFF'),(2227,'03 pos','ticket-order-format',''),(2228,'00 sales','ticket-order-format',''),(2230,'03 pos','label-order-format',''),(2231,'00 sales','label-order-format',''),(2242,'00 sales','keyboard-orders',''),(2252,'01 packet','keyboard-orders',''),(2262,'02 selfservice','keyboard-orders',''),(2272,'03 pos','keyboard-orders',''),(2274,'00 sales','unhang-tickets-time','0'),(2275,'00 sales','user-pricerates','ON'),(2276,'00 sales','allow-customer-order','ON'),(2320,'00 sales','weight-min','0,000'),(2321,'00 sales','amount-min','0,00'),(2322,'00 sales','pricerate','0'),(2336,'00 sales','cancel-hanged','0'),(2337,'00 sales','cancel-hanged-time','10'),(2338,'00 sales','group-order-tickets','OFF'),(2339,'00 sales','group-order-as-invoice','ON'),(2340,'00 sales','group-parked-tickets','OFF'),(2341,'00 sales','group-parked-as-invoice','ON'),(2342,'00 sales','drawer-with-password','OFF'),(2343,'00 sales','drawer-key-on-vendor-menu','OFF'),(2344,'00 sales','allow-ticket-check','ON'),(2345,'00 sales','max-lines-99','OFF'),(2346,'00 sales','save-code128','0'),(2393,'00 sales','change-price0','ON'),(2394,'00 sales','change-saleform','OFF');
/*!40000 ALTER TABLE `localworkingmodes` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalWMInserted` AFTER INSERT ON `LocalWorkingModes`
FOR EACH ROW BEGIN
  IF(NEW.`WorkingMode`=(SELECT `ParamValue` FROM `LocalWorkingModes` WHERE `ParamName`="working-mode" AND `WorkingMode`="")) THEN
    SET @trx := (SELECT trx_new("BUTTON_UPDATE", "CORE"));
    SET @ret := (SELECT trx_set_param(@trx, "WORKINGMODE", NEW.`WorkingMode`));
    SET @ret := (SELECT trx_set_param(@trx, "PARAMNAME", CONCAT("WORKINGMODES_", NEW.`ParamName`)));
    SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", NEW.`ParamValue`));
    SET @ret := (SELECT trx_send(@trx));
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalWMUpdated` AFTER UPDATE ON `LocalWorkingModes`
FOR EACH ROW BEGIN
  IF(NEW.`WorkingMode`=(SELECT `ParamValue` FROM `LocalWorkingModes` WHERE `ParamName`="working-mode" AND `WorkingMode`="")) THEN
    IF(OLD.`ParamValue`!=NEW.`ParamValue`) THEN
      SET @trx := (SELECT trx_new("BUTTON_UPDATE", "CORE"));
      SET @ret := (SELECT trx_set_param(@trx, "WORKINGMODE", NEW.`WorkingMode`));
      SET @ret := (SELECT trx_set_param(@trx, "PARAMNAME", CONCAT("WORKINGMODES_", NEW.`ParamName`)));
      SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", NEW.`ParamValue`));
      SET @ret := (SELECT trx_send(@trx));
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalWMDeleted` AFTER DELETE ON `LocalWorkingModes`
FOR EACH ROW BEGIN
  IF(OLD.`WorkingMode`=(SELECT `ParamValue` FROM `LocalWorkingModes` WHERE `ParamName`="working-mode" AND `WorkingMode`="")) THEN
    SET @trx := (SELECT trx_new("BUTTON_UPDATE", "CORE"));
    SET @ret := (SELECT trx_set_param(@trx, "WORKINGMODE", OLD.`WorkingMode`));
    SET @ret := (SELECT trx_set_param(@trx, "PARAMNAME", CONCAT("WORKINGMODES_", OLD.`ParamName`)));
    SET @ret := (SELECT trx_set_param(@trx, "PARAMVALUE", ""));
    SET @ret := (SELECT trx_send(@trx));
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `ltickets`
--

DROP TABLE IF EXISTS `ltickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `ltickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `IdHTicket` int(11) DEFAULT NULL,
  `NetStat` char(1) DEFAULT 'N',
  `lineType` smallint(6) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `SaleFormEx` char(1) DEFAULT '-',
  `Item` int(11) DEFAULT NULL,
  `Name` varchar(127) NOT NULL DEFAULT '',
  `Section` decimal(3,0) NOT NULL DEFAULT '0',
  `Family` int(11) DEFAULT '0',
  `ItemAccess` int(11) DEFAULT '0',
  `LineScale` int(11) NOT NULL DEFAULT '0',
  `LineVendor` int(11) NOT NULL DEFAULT '0',
  `LineDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `SourceKey` char(1) NOT NULL DEFAULT '0',
  `PriceChanged` tinyint(1) NOT NULL DEFAULT '0',
  `VATCode` decimal(1,0) unsigned NOT NULL DEFAULT '0',
  `VATPercent` decimal(4,2) unsigned NOT NULL DEFAULT '0.00',
  `Weight` decimal(7,3) DEFAULT NULL,
  `Tare` decimal(7,3) DEFAULT '0.000',
  `TTare` char(1) DEFAULT '',
  `lineDiscount` decimal(4,2) DEFAULT NULL,
  `Price` decimal(8,2) DEFAULT NULL,
  `Amount` decimal(10,2) DEFAULT NULL,
  `Points` int(11) NOT NULL DEFAULT '0',
  `BovineCard` smallint(3) DEFAULT NULL,
  `EarTag` varchar(30) DEFAULT NULL,
  `BovineLot` varchar(20) DEFAULT NULL,
  `FreezingDate` date DEFAULT NULL,
  `TicketDiscount` decimal(4,2) DEFAULT '0.00',
  `PrintScale` int(11) NOT NULL DEFAULT '0',
  `PrintDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `PrintTicketNumber` int(11) NOT NULL DEFAULT '0',
  `PrinttckType` char(1) DEFAULT '',
  `ReloadScale` int(11) NOT NULL DEFAULT '0',
  `ReloadDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadTicketNumber` int(11) NOT NULL DEFAULT '0',
  `ReloadtckType` char(1) DEFAULT '',
  `ReloadTo` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  KEY `IdHTicket` (`IdHTicket`),
  CONSTRAINT `ltickets_ibfk_1` FOREIGN KEY (`IdHTicket`) REFERENCES `htickets` (`Id`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=29 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `ltickets`
--

LOCK TABLES `ltickets` WRITE;
/*!40000 ALTER TABLE `ltickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `ltickets` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LTicketsInsertB` BEFORE INSERT ON `LTickets`
FOR EACH ROW BEGIN

  IF ((NEW.`NetStat`='N') AND  (NEW.`lineType`=1)) THEN
    IF ((SELECT Status FROM `HTickets` WHERE `HTickets`.`Id`=NEW.`IdHTicket`)='O') THEN
      SET NEW.`LineDateTime`=NOW();
      IF ((SELECT IFNULL(`ParamValue`,0) FROM `CfgSystem` WHERE `ParamKey`='loyalty-items-extra-points')!=0) THEN
        SET NEW.`Points`=IFNULL((SELECT
          IF(`ItemsExtraPoints`.`FromQuantity`,

            IF(`ItemsExtraPoints`.`Quantity`,

              IF(`ItemsExtraPoints`.`Proportional`,

                IF(NEW.`Weight`>=`ItemsExtraPoints`.`FromQuantity`,(((NEW.`Weight`-`ItemsExtraPoints`.`FromQuantity`)/`ItemsExtraPoints`.`Quantity`) +1)*`ItemsExtraPoints`.`Points`, 0),

                  IF(NEW.`Weight`>=`ItemsExtraPoints`.`FromQuantity`,(((NEW.`Weight`-`ItemsExtraPoints`.`FromQuantity`)/`ItemsExtraPoints`.`Quantity`) DIV 1 +1)*`ItemsExtraPoints`.`Points`, 0)),

              IF(NEW.`Weight`>=`ItemsExtraPoints`.`FromQuantity`,`ItemsExtraPoints`.`Points`,0)),

            IF(`ItemsExtraPoints`.`Quantity`,

              IF(`ItemsExtraPoints`.`Proportional`,

                (NEW.`Weight`/`ItemsExtraPoints`.`Quantity`)*`ItemsExtraPoints`.`Points`,

                  IF(NEW.`Weight`>=`ItemsExtraPoints`.`Quantity`,(((NEW.`Weight`-`ItemsExtraPoints`.`Quantity`)/`ItemsExtraPoints`.`Quantity`) DIV 1 +1)*`ItemsExtraPoints`.`Points`, 0)),

              `ItemsExtraPoints`.`Points`))
        FROM `ItemsExtraPoints` WHERE `ItemsExtraPoints`.`Code`=NEW.`Item`), 0);
      END IF;

    ELSE
      UPDATE `PCScaleErrorTag` SET x=1;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LTicketsInserted` AFTER INSERT ON `LTickets`
FOR EACH ROW BEGIN

  IF(NEW.`NetStat`='N') THEN
    SET @wgt = IF(NEW.`SaleForm`=1,NEW.`Weight`,0);
    SET @units = IF(NEW.`SaleForm`=0,NEW.`Weight`,0);
    CASE(NEW.`lineType`)

      WHEN 0 THEN
        UPDATE `HTickets` SET `NumLines`=`NumLines`+1, `CanceledLines`=`CanceledLines`+1, `CanceledAmount`=`CanceledAmount`+NEW.`Amount`,
          `DatTim`=NOW(), `SectionItem`=NEW.`Section` WHERE `HTickets`.`Id`=NEW.`IdHTicket` AND `HTickets`.`Status`='O';

      WHEN 1 THEN
        SELECT `HTickets`.`OrderNumber`,`HTickets`.`SectionItem` INTO @norder, @secorder FROM `HTickets` WHERE `HTickets`.`Id`=NEW.`IdHTicket` AND `HTickets`.`Status`='O';
        UPDATE `HTickets` SET `NumLines`=`NumLines`+1, `AddedLines`=`AddedLines`+1, `GrossAmount`=`GrossAmount`+NEW.`Amount`,
          `NetAmount`=CAST(`GrossAmount`*((100-`tckDiscount`)/100) AS DECIMAL(10,2)),
          `NegativeLines`=`NegativeLines`+IF(NEW.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`+IF(NEW.`Price`<0,-NEW.`Amount`,0),
          `TotalWeight`=`TotalWeight`+IF(NEW.`Price`<0,0,@wgt), `TotalUnits`=`TotalUnits`+IF(NEW.`Price`<0,0,@units),
          `HTickets`.`PointsExtra`=`HTickets`.`PointsExtra`+IF(NEW.`Price`>0,NEW.`Points`,0),
          `DatTim`=NOW(), `SectionItem`=NEW.`Section` WHERE `HTickets`.`Id`=NEW.`IdHTicket` AND `HTickets`.`Status`='O';
        IF(@norder!=0) THEN 
          INSERT INTO `LCustomerOrders` (`OrderNumber`, `Section`, `Item`, `Price`,`Status`,`Weight`,`Units`,`Amount`) VALUES (@norder, @secorder, NEW.`Item`, NEW.`Price`, 3, @wgt, @units,NEW.`Amount`)
            ON DUPLICATE KEY UPDATE `LCustomerOrders`.`Weight`=`LCustomerOrders`.`Weight`+@wgt, `LCustomerOrders`.`Units`=`LCustomerOrders`.`Units`+@units,
            `LCustomerOrders`.`Amount`=`LCustomerOrders`.`Amount`+VALUES(`Amount`),
            `LCustomerOrders`.`Status`=((CAST(Status AS UNSIGNED) & CAST(b'00000010' AS UNSIGNED)) | CAST(b'00000001' AS UNSIGNED));
         END IF;

      WHEN 3 THEN
        IF(NEW.`SaleForm`=0) THEN
          UPDATE `HTickets` SET `TotalWeight`=`TotalWeight`+@wgt, `TotalUnits`=`TotalUnits`+@units, `PointsChanged`=`PointsChanged`+NEW.`Price`, `DatTim`=NOW()
            WHERE `HTickets`.`Id`=NEW.`IdHTicket` AND `HTickets`.`Status`='O';
        END IF;

      ELSE BEGIN END;
    END CASE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LTicketsUpdated` AFTER UPDATE ON `LTickets`
FOR EACH ROW BEGIN
  IF((NEW.`IdHTicket`!=OLD.`IdHTicket`) OR (NEW.`Amount`!=OLD.`Amount`) OR (NEW.`VATCode`!=OLD.`VATCode`) OR (NEW.`VATPercent`!=OLD.`VATPercent`)
    OR (NEW.`SaleForm`!=OLD.`SaleForm`) OR (NEW.`Weight`!=OLD.`Weight`) OR (NEW.`Price`!=OLD.`Price`) OR (NEW.`lineDiscount`!=OLD.`lineDiscount`)) THEN
    CALL `manipulate_protected_field`;
  END IF;
  IF(NEW.`lineType`!=OLD.`lineType`) THEN
    SET @wgt = IF(OLD.`SaleForm`=1,OLD.`Weight`,0);
    SET @units = IF(NEW.`SaleForm`=0,NEW.`Weight`,0);
    SET @stck = ' ';
    SET @norder = 0;
    SELECT `HTickets`.`Status`, `HTickets`.`OrderNumber`INTO @stck,@norder FROM `HTickets` WHERE `HTickets`.`Id`=OLD.`IdHTicket`;
    IF(@stck='O') THEN
      CASE(NEW.`lineType`)

        WHEN 0 THEN
          IF(OLD.`lineType`=1) THEN
            UPDATE `HTickets` SET `AddedLines`=`AddedLines`-1, `GrossAmount`=`GrossAmount`-OLD.`Amount`,
              `NetAmount`=CAST(`GrossAmount`*((100-`tckDiscount`)/100) AS DECIMAL(10,2)),
              `NegativeLines`=`NegativeLines`-IF(OLD.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`-IF(OLD.`Price`<0,-OLD.`Amount`,0),
              `CanceledLines`=`CanceledLines`+1, `CanceledAmount` = `CanceledAmount`+OLD.`Amount`,
              `TotalWeight`=`TotalWeight`-IF(OLD.`Price`<0,0,@wgt), `TotalUnits`=`TotalUnits`-IF(NEW.`Price`<0,0,@units),
              `HTickets`.`PointsExtra`=`HTickets`.`PointsExtra`-IF(OLD.`Price`>0,OLD.`Points`,0),
              `DatTim`=NOW()
            WHERE `HTickets`.`Id`=OLD.`IdHTicket`;
            IF(@norder!=0) THEN 
              UPDATE `LCustomerOrders` SET `LCustomerOrders`.`Weight`=`LCustomerOrders`.`Weight`-@wgt, `LCustomerOrders`.`Units`=`LCustomerOrders`.`Units`-@units,
                `LCustomerOrders`.`Amount`=`LCustomerOrders`.`Amount`-OLD.`Amount`,
                `LCustomerOrders`.`Status`=IF((`LCustomerOrders`.`Status`=1 AND `LCustomerOrders`.`Weight`=0 AND `LCustomerOrders`.`Units`=0), 0, `LCustomerOrders`.`Status`)
              WHERE `LCustomerOrders`.`OrderNumber`=@norder AND `LCustomerOrders`.`Item`=OLD.`Item`;
            END IF;
          END IF;

        WHEN 1 THEN
          IF(OLD.`lineType`=0) THEN
            UPDATE `HTickets` SET `AddedLines`=`AddedLines`+1, `GrossAmount`=`GrossAmount`+OLD.`Amount`,
              `NetAmount`=CAST(`GrossAmount`*((100-`tckDiscount`)/100) AS DECIMAL(10,2)),
              `NegativeLines`=`NegativeLines`+IF(OLD.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`+IF(OLD.`Price`<0,-OLD.`Amount`,0),
              `CanceledLines`=`CanceledLines`-1, `CanceledAmount` = `CanceledAmount`-OLD.`Amount`,
              `TotalWeight`=`TotalWeight`+IF(OLD.`Price`<0,0,@wgt), `TotalUnits`=`TotalUnits`+IF(NEW.`Price`<0,0,@units),
              `HTickets`.`PointsExtra`=`HTickets`.`PointsExtra`+IF(OLD.`Price`>0,OLD.`Points`,0),
              `DatTim`=NOW()
            WHERE `HTickets`.`Id`=OLD.`IdHTicket`;
            IF(@norder!=0) THEN 
              UPDATE `LCustomerOrders` SET `LCustomerOrders`.`Weight`=`LCustomerOrders`.`Weight`+@wgt, `LCustomerOrders`.`Units`=`LCustomerOrders`.`Units`+@units,
                `LCustomerOrders`.`Amount`=`LCustomerOrders`.`Amount`+OLD.`Amount`,
                `LCustomerOrders`.`Status`=IF(`LCustomerOrders`.`Status`=0, 1, `LCustomerOrders`.`Status`)
              WHERE `LCustomerOrders`.`OrderNumber`=@norder AND `LCustomerOrders`.`Item`=OLD.`Item`;
            END IF;
          END IF;

        WHEN 2 THEN
          IF((NEW.`SaleForm`=0) AND (OLD.`lineType`=3)) THEN
            UPDATE `HTickets` SET `TotalWeight`=`TotalWeight`-@wgt, `TotalUnits`=`TotalUnits`-@units, `PointsChanged`=`PointsChanged`-OLD.`Price`, `DatTim`=NOW() WHERE `HTickets`.`Id`=OLD.`IdHTicket`;
          END IF;

        WHEN 3 THEN
          IF((NEW.`SaleForm`=0) AND (OLD.`lineType`=2)) THEN
            UPDATE `HTickets` SET `TotalWeight`=`TotalWeight`+@wgt, `TotalUnits`=`TotalUnits`+@units, `PointsChanged`=`PointsChanged`+OLD.`Price`, `DatTim`=NOW() WHERE `HTickets`.`Id`=OLD.`IdHTicket`;
          END IF;

        ELSE BEGIN END;
      END CASE;
    END IF;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `netstatus`
--

DROP TABLE IF EXISTS `netstatus`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `netstatus` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `NetworkName` varchar(50) NOT NULL DEFAULT '',
  `NetworkPrivateName` varchar(80) NOT NULL DEFAULT '',
  `MCast_port` decimal(4,0) NOT NULL DEFAULT '2335',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `netstatus`
--

LOCK TABLES `netstatus` WRITE;
/*!40000 ALTER TABLE `netstatus` DISABLE KEYS */;
INSERT INTO `netstatus` VALUES (1,'','',2335);
/*!40000 ALTER TABLE `netstatus` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `numberingtickets`
--

DROP TABLE IF EXISTS `numberingtickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `numberingtickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` binary(1) DEFAULT 'T',
  `Numbering` smallint(6) DEFAULT '0',
  `NumberingLocal` tinyint(1) DEFAULT '0',
  `StartNumber` int(10) unsigned DEFAULT '1',
  `MaximumNumber` int(10) unsigned DEFAULT '9999',
  `MaximumInvoice` int(10) unsigned DEFAULT '9999',
  `ResetNumber` tinyint(1) DEFAULT '0',
  `SaveMode` char(1) DEFAULT '',
  `ClearStart` tinyint(1) DEFAULT '0',
  `ClearStartDays` smallint(6) DEFAULT '0',
  `ClearStartResetNumber` tinyint(1) DEFAULT '0',
  `ClearStartDaily` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Type` (`Type`)
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `numberingtickets`
--

LOCK TABLES `numberingtickets` WRITE;
/*!40000 ALTER TABLE `numberingtickets` DISABLE KEYS */;
INSERT INTO `numberingtickets` VALUES (10,'T',1,0,1,99999,99999,1,'D',1,366,0,0),(11,'P',1,0,1,99999,9999,1,'D',0,0,0,0),(12,'S',1,0,1,99999,9999,1,'D',0,0,0,0),(13,'C',1,0,8000,99999,9999,1,'D',0,0,0,0),(14,'D',1,0,1,9999,9999,1,'D',0,0,0,0),(15,'F',1,0,1,9999,9999,1,'D',0,0,0,0);
/*!40000 ALTER TABLE `numberingtickets` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `numinvoices`
--

DROP TABLE IF EXISTS `numinvoices`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `numinvoices` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `MaxNumber` int(11) DEFAULT NULL,
  `Number` int(11) DEFAULT NULL,
  `Serial` varchar(16) NOT NULL DEFAULT '',
  `RecNumber` int(11) DEFAULT NULL,
  `RecSerial` varchar(16) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `numinvoices`
--

LOCK TABLES `numinvoices` WRITE;
/*!40000 ALTER TABLE `numinvoices` DISABLE KEYS */;
INSERT INTO `numinvoices` VALUES (1,9999,0,'2016/',0,'R');
/*!40000 ALTER TABLE `numinvoices` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `numtickets`
--

DROP TABLE IF EXISTS `numtickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `numtickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` binary(1) DEFAULT 'T',
  `Identifier` varchar(16) NOT NULL DEFAULT '',
  `Number` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Identifier` (`Identifier`,`Type`)
) ENGINE=InnoDB AUTO_INCREMENT=18 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `numtickets`
--

LOCK TABLES `numtickets` WRITE;
/*!40000 ALTER TABLE `numtickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `numtickets` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `nutrition`
--

DROP TABLE IF EXISTS `nutrition`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `nutrition` (
  `card` int(11) NOT NULL,
  `portion_weight` float DEFAULT '0',
  `energy_kj_100` float DEFAULT '0',
  `energy_kj_portion` float DEFAULT '0',
  `energy_kcal_100` float DEFAULT '0',
  `energy_kcal_portion` float DEFAULT '0',
  `fat_100` float DEFAULT '0',
  `fat_portion` float DEFAULT '0',
  `saturates_100` float DEFAULT '0',
  `saturates_portion` float DEFAULT '0',
  `mono_unsaturates_100` float DEFAULT '0',
  `mono_unsaturates_portion` float DEFAULT '0',
  `polyunsaturates_100` float DEFAULT '0',
  `polyunsaturates_portion` float DEFAULT '0',
  `carbohydrate_100` float DEFAULT '0',
  `carbohydrate_portion` float DEFAULT '0',
  `sugars_100` float DEFAULT '0',
  `sugars_portion` float DEFAULT '0',
  `polyols_100` float DEFAULT '0',
  `polyols_portion` float DEFAULT '0',
  `starch_100` float DEFAULT '0',
  `starch_portion` float DEFAULT '0',
  `fibre_100` float DEFAULT '0',
  `fibre_portion` float DEFAULT '0',
  `protein_100` float DEFAULT '0',
  `protein_portion` float DEFAULT '0',
  `salt_100` float DEFAULT '0',
  `salt_portion` float DEFAULT '0',
  PRIMARY KEY (`card`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `nutrition`
--

LOCK TABLES `nutrition` WRITE;
/*!40000 ALTER TABLE `nutrition` DISABLE KEYS */;
INSERT INTO `nutrition` VALUES (1,10,10,5,20,10,200,100,300,150,400,200,500,250,44,55,555,66,22,33,511,55,11,6,22,12,33,15);
/*!40000 ALTER TABLE `nutrition` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `packetlabels`
--

DROP TABLE IF EXISTS `packetlabels`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `packetlabels` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Captured` char(1) DEFAULT '',
  `CapturedDateTime` datetime NOT NULL DEFAULT '1970-01-01 00:00:00',
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `lineType` smallint(6) DEFAULT NULL,
  `Item` int(11) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `Weight` decimal(7,3) DEFAULT NULL,
  `Tare` decimal(7,3) DEFAULT '0.000',
  `TTare` char(1) DEFAULT '',
  `Price` decimal(8,2) DEFAULT NULL,
  `lineDiscount` decimal(4,2) DEFAULT NULL,
  `Amount` decimal(10,2) DEFAULT NULL,
  `PacketExt` char(1) NOT NULL DEFAULT '-',
  `BovineCard` int(11) DEFAULT NULL,
  `EarTag` varchar(30) DEFAULT NULL,
  `BovineLot` varchar(20) DEFAULT NULL,
  `FreezingDate` date DEFAULT NULL,
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `DatTim` datetime DEFAULT NULL,
  `EAN13` varchar(13) NOT NULL DEFAULT '',
  `Code128` varchar(32) NOT NULL DEFAULT '',
  `Lot` varchar(20) NOT NULL DEFAULT '',
  `Ordinal` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=23 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `packetlabels`
--

LOCK TABLES `packetlabels` WRITE;
/*!40000 ALTER TABLE `packetlabels` DISABLE KEYS */;
/*!40000 ALTER TABLE `packetlabels` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `paymethods`
--

DROP TABLE IF EXISTS `paymethods`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `paymethods` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Identifier` tinyint(3) unsigned NOT NULL,
  `Name` varchar(64) NOT NULL,
  `InUse` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Identifier` (`Identifier`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `paymethods`
--

LOCK TABLES `paymethods` WRITE;
/*!40000 ALTER TABLE `paymethods` DISABLE KEYS */;
INSERT INTO `paymethods` VALUES (1,0,'EFECTIVO',1),(2,1,'T.CREDITO',1),(3,2,'CHEQUE',1),(4,3,'PAYMETHOD3',0),(5,4,'PAYMETHOD4',0),(6,5,'PAYMETHOD5',0),(7,6,'PAYMETHOD6',0);
/*!40000 ALTER TABLE `paymethods` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `paytickets`
--

DROP TABLE IF EXISTS `paytickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `paytickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `IdHTicket` int(11) DEFAULT NULL,
  `PayMethod` int(11) DEFAULT NULL,
  `Delivered` decimal(10,2) DEFAULT '0.00',
  PRIMARY KEY (`Id`),
  KEY `IdHTicket` (`IdHTicket`),
  CONSTRAINT `paytickets_ibfk_1` FOREIGN KEY (`IdHTicket`) REFERENCES `htickets` (`Id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `paytickets`
--

LOCK TABLES `paytickets` WRITE;
/*!40000 ALTER TABLE `paytickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `paytickets` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pktfishery`
--

DROP TABLE IF EXISTS `pktfishery`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `pktfishery` (
  `Code` int(11) NOT NULL,
  `SCName` varchar(127) NOT NULL,
  `Lot` varchar(20) NOT NULL,
  `Date` date DEFAULT NULL,
  `Production` varchar(127) NOT NULL,
  `Catched` varchar(127) NOT NULL,
  `Presentation` varchar(127) NOT NULL,
  `Size` decimal(1,0) DEFAULT NULL,
  `Freshness` varchar(10) NOT NULL,
  `Issuer` varchar(255) NOT NULL,
  `Descriptive` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`Code`),
  CONSTRAINT `pktfishery_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `pktfishery`
--

LOCK TABLES `pktfishery` WRITE;
/*!40000 ALTER TABLE `pktfishery` DISABLE KEYS */;
/*!40000 ALTER TABLE `pktfishery` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pktgeneric`
--

DROP TABLE IF EXISTS `pktgeneric`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `pktgeneric` (
  `NumCard` smallint(3) NOT NULL,
  `Lot` varchar(20) NOT NULL,
  `OriginLand` smallint(6) DEFAULT NULL,
  `Region` varchar(255) NOT NULL,
  `SanitaryReg1` varchar(30) DEFAULT NULL,
  `Date1` date DEFAULT NULL,
  `SanitaryReg2` varchar(30) DEFAULT NULL,
  `Date2` date DEFAULT NULL,
  `Descriptive` varchar(255) DEFAULT NULL,
  `Text1` varchar(255) DEFAULT NULL,
  `Text2` varchar(255) DEFAULT NULL,
  `Text3` varchar(255) DEFAULT NULL,
  `Text4` varchar(255) DEFAULT NULL,
  `Text5` varchar(255) DEFAULT NULL,
  `Text6` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`NumCard`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `pktgeneric`
--

LOCK TABLES `pktgeneric` WRITE;
/*!40000 ALTER TABLE `pktgeneric` DISABLE KEYS */;
/*!40000 ALTER TABLE `pktgeneric` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pktvegetables`
--

DROP TABLE IF EXISTS `pktvegetables`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `pktvegetables` (
  `Code` int(11) NOT NULL,
  `Variety` varchar(127) NOT NULL,
  `Lot` varchar(20) NOT NULL,
  `OriginLand` smallint(6) DEFAULT NULL,
  `Region` varchar(255) NOT NULL,
  `Quality` varchar(127) NOT NULL,
  `Size` decimal(2,0) DEFAULT NULL,
  `Issuer` varchar(255) NOT NULL,
  `Descriptive` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`Code`),
  CONSTRAINT `pktvegetables_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `pktvegetables`
--

LOCK TABLES `pktvegetables` WRITE;
/*!40000 ALTER TABLE `pktvegetables` DISABLE KEYS */;
/*!40000 ALTER TABLE `pktvegetables` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `playlists`
--

DROP TABLE IF EXISTS `playlists`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `playlists` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Name` varchar(31) DEFAULT NULL,
  `Type` smallint(6) DEFAULT '0',
  `PlayIndex` int(11) DEFAULT NULL,
  `FileName` varchar(127) NOT NULL,
  `ExposureTime` smallint(6) DEFAULT '0',
  `TransitionTime` smallint(6) DEFAULT '0',
  `TransitionType` smallint(6) DEFAULT '0',
  `StartDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `EndDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `Status` char(1) DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Name` (`Name`,`PlayIndex`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `playlists`
--

LOCK TABLES `playlists` WRITE;
/*!40000 ALTER TABLE `playlists` DISABLE KEYS */;
INSERT INTO `playlists` VALUES (1,'a',0,1,'beef2_625-348.png',5,0,0,'0000-00-00 00:00:00','0000-00-00 00:00:00',''),(2,'a',0,2,'beef 625x348.png',5,0,0,'0000-00-00 00:00:00','0000-00-00 00:00:00',''),(3,'a',0,3,'fish 625x348.png',5,0,0,'0000-00-00 00:00:00','0000-00-00 00:00:00',''),(4,'a',0,4,'fruits2 625x348.png',5,0,0,'0000-00-00 00:00:00','0000-00-00 00:00:00',''),(5,'a',0,5,'gummi-bears 625x348.png',5,0,0,'0000-00-00 00:00:00','0000-00-00 00:00:00',''),(6,'1',1,1,'Maravillosa Naturaleza HD (WONDERFUL NATURE).avi',5,0,0,'0000-00-00 00:00:00','0000-00-00 00:00:00','');
/*!40000 ALTER TABLE `playlists` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `plukeys`
--

DROP TABLE IF EXISTS `plukeys`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `plukeys` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Section` int(11) DEFAULT NULL,
  `Board` int(11) DEFAULT NULL,
  `PLUKey` int(11) DEFAULT NULL,
  `Code` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Section` (`Section`,`Board`,`PLUKey`),
  KEY `Code` (`Code`)
) ENGINE=InnoDB AUTO_INCREMENT=1599 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `plukeys`
--

LOCK TABLES `plukeys` WRITE;
/*!40000 ALTER TABLE `plukeys` DISABLE KEYS */;
/*!40000 ALTER TABLE `plukeys` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`localhost`*/ /*!50003 TRIGGER `PLUKeysInserted` AFTER INSERT ON `PLUKeys`
FOR EACH ROW BEGIN
  IF (NEW.`Section`=0 AND NEW.`Board`=GLOBAL_GET("current_mupboard")) THEN
    INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (0, NEW.`Board`, NEW.`PLUKey`, NEW.`Code`)
      ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;
  ELSEIF (NEW.`Section`=GLOBAL_GET("current_group")) THEN
    CASE GLOBAL_GET("PLUload-func")

      WHEN 0 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (NEW.`Section`, NEW.`Board`, NEW.`PLUKey`, NEW.`Code`)
        ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;

      WHEN 1 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Section`, `Family`, NEW.`PLUKey`, NEW.`Code` FROM `Items` WHERE `Items`.`Code`=NEW.`Code`
        ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;

      WHEN 2 THEN BEGIN END;




      ELSE BEGIN END;
    END CASE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`localhost`*/ /*!50003 TRIGGER `PLUKeysUpdated` AFTER UPDATE ON `PLUKeys`
FOR EACH ROW BEGIN
  SET @mygroup := GLOBAL_GET("current_group");
  SET @mupboard := GLOBAL_GET("current_mupboard");
  IF (OLD.`Section`=0 AND OLD.`Board`=@mupboard) THEN
      INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (0, @mupboard, OLD.`PLUKey`, 0)
      ON DUPLICATE KEY UPDATE `Code`=0;
  ELSEIF (OLD.`Section`=@mygroup) THEN
    CASE GLOBAL_GET("PLUload-func")

      WHEN 0 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (OLD.`Section`, OLD.`Board`, OLD.`PLUKey`, 0)
        ON DUPLICATE KEY UPDATE `Code`=0;

      WHEN 1 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT OLD.`Section`, `Family`, OLD.`PLUKey`, 0 FROM `Items` WHERE `Items`.`Code`=OLD.`Code`
        ON DUPLICATE KEY UPDATE `Code`=0;

      WHEN 2 THEN BEGIN END;




      ELSE BEGIN END;
   END CASE;
  END IF;
  IF (NEW.`Section`=0 AND NEW.`Board`=@mupboard) THEN
    INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (0, @mupboard, NEW.`PLUKey`, NEW.`Code`)
          ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;
  ELSEIF (NEW.`Section`=@mygroup) THEN
    CASE GLOBAL_GET("PLUload-func")

      WHEN 0 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (NEW.`Section`, NEW.`Board`, NEW.`PLUKey`, NEW.`Code`)
        ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;

      WHEN 1 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT NEW.`Section`, `Family`, NEW.`PLUKey`, NEW.`Code` FROM `Items` WHERE `Items`.`Code`=NEW.`Code`
        ON DUPLICATE KEY UPDATE `Code`=NEW.`Code`;

      WHEN 2 THEN BEGIN END;




      ELSE BEGIN END;
    END CASE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`localhost`*/ /*!50003 TRIGGER `PLUKeysDeleted` AFTER DELETE ON `PLUKeys`
FOR EACH ROW BEGIN
  IF (OLD.`Section`=0 AND OLD.`Board`=GLOBAL_GET("current_mupboard")) THEN
    INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (0, OLD.`Board`, OLD.`PLUKey`, 0)
    ON DUPLICATE KEY UPDATE `Code`=0;
  ELSEIF (OLD.`Section`=GLOBAL_GET("current_group")) THEN
    CASE GLOBAL_GET("PLUload-func")

      WHEN 0 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) VALUES (OLD.`Section`, OLD.`Board`, OLD.`PLUKey`, 0)
        ON DUPLICATE KEY UPDATE `Code`=0;

      WHEN 1 THEN
        INSERT INTO `LocalPLUKeysChanges` (`Section`,`Board`,`PLUKey`,`Code`) SELECT OLD.`Section`, `Family`, OLD.`PLUKey`, 0 FROM `Items` WHERE `Items`.`Code`=OLD.`Code`
        ON DUPLICATE KEY UPDATE `Code`=0;

      WHEN 2 THEN BEGIN END;




      ELSE BEGIN END;
    END CASE;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `pricerates`
--

DROP TABLE IF EXISTS `pricerates`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `pricerates` (
  `Code` decimal(1,0) NOT NULL,
  `Name` varchar(64) NOT NULL DEFAULT '',
  PRIMARY KEY (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `pricerates`
--

LOCK TABLES `pricerates` WRITE;
/*!40000 ALTER TABLE `pricerates` DISABLE KEYS */;
INSERT INTO `pricerates` VALUES (0,''),(1,'Tar. 1'),(2,'Tar. 2'),(3,'Tar. 3'),(4,'Tar. 4'),(5,'Tar. 5'),(6,'Tar. 6'),(7,'Tar. 7'),(8,'Tar. 8'),(9,'Tar. 9');
/*!40000 ALTER TABLE `pricerates` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `qlevels`
--

DROP TABLE IF EXISTS `qlevels`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `qlevels` (
  `Code` varchar(16) NOT NULL,
  `Icon` varchar(64) NOT NULL,
  `Description` varchar(127) NOT NULL,
  PRIMARY KEY (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `qlevels`
--

LOCK TABLES `qlevels` WRITE;
/*!40000 ALTER TABLE `qlevels` DISABLE KEYS */;
INSERT INTO `qlevels` VALUES ('1','F_001.png','');
/*!40000 ALTER TABLE `qlevels` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `qrfields`
--

DROP TABLE IF EXISTS `qrfields`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `qrfields` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `SourceTable` char(1) NOT NULL DEFAULT 'H',
  `Name` varchar(32) NOT NULL DEFAULT '',
  `FieldIndex` decimal(2,0) NOT NULL DEFAULT '1',
  `FieldType` char(1) NOT NULL DEFAULT 'F',
  `Size` decimal(2,0) NOT NULL DEFAULT '1',
  `Format` varchar(127) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `qrf_idx` (`SourceTable`,`FieldIndex`)
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `qrfields`
--

LOCK TABLES `qrfields` WRITE;
/*!40000 ALTER TABLE `qrfields` DISABLE KEYS */;
INSERT INTO `qrfields` VALUES (1,'H','[]',1,'F',0,'123'),(2,'H','Number',2,'N',6,'0FILLED'),(3,'H','DatTim',3,'D',0,'%y%m%d%H%i%s'),(4,'H','AddedLines',4,'N',2,'0FILLED'),(5,'H','NetAmount',5,'M',6,'0FILLED DPNONE'),(6,'L','[]',1,'F',0,'1'),(7,'L','Item',2,'N',5,'0FILLED'),(8,'L','Weight',3,'Q',5,'0FILLED DPNONE'),(9,'L','Price',4,'M',5,'0FILLED DPNONE'),(10,'L','Amount',5,'M',5,'0FILLED DPNONE'),(11,'','[]',1,'',0,'check128');
/*!40000 ALTER TABLE `qrfields` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `quickkeys`
--

DROP TABLE IF EXISTS `quickkeys`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `quickkeys` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Label` varchar(64) DEFAULT '',
  `Vendor` int(11) DEFAULT '0',
  `EAN13` char(13) NOT NULL DEFAULT '',
  `Code128` decimal(1,0) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `quickkeys`
--

LOCK TABLES `quickkeys` WRITE;
/*!40000 ALTER TABLE `quickkeys` DISABLE KEYS */;
/*!40000 ALTER TABLE `quickkeys` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `scales`
--

DROP TABLE IF EXISTS `scales`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `scales` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `MAC` varchar(17) NOT NULL,
  `IP` varchar(16) NOT NULL,
  `Number` int(11) NOT NULL DEFAULT '0',
  `Counter` int(11) NOT NULL DEFAULT '0',
  `Name` varchar(64) NOT NULL DEFAULT '',
  `Type` char(1) DEFAULT NULL,
  `AppVersion` varchar(16) DEFAULT '',
  `ScaleVersion` varchar(32) NOT NULL DEFAULT '',
  `SerialNumber` varchar(10) DEFAULT '',
  `System` varchar(32) DEFAULT '',
  `Keyboard` varchar(128) NOT NULL DEFAULT '',
  `WorkingMode` varchar(128) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `MAC` (`MAC`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `scales`
--

LOCK TABLES `scales` WRITE;
/*!40000 ALTER TABLE `scales` DISABLE KEYS */;
INSERT INTO `scales` VALUES (1,'00:30:18:C6:57:B3','192.168.85.143',1,4,'','A','2.20.3003','','0','Ubuntu_precise','1024x768/NeutralStyle/sales_TurnFull','00 sales');
/*!40000 ALTER TABLE `scales` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = latin1 */ ;
/*!50003 SET character_set_results = latin1 */ ;
/*!50003 SET collation_connection  = latin1_swedish_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `JoinScale` AFTER INSERT ON `Scales`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `RemoveScale`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = latin1 */ ;
/*!50003 SET character_set_results = latin1 */ ;
/*!50003 SET collation_connection  = latin1_swedish_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `ChangeScale` AFTER UPDATE ON `Scales`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `RemoveScale`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = latin1 */ ;
/*!50003 SET character_set_results = latin1 */ ;
/*!50003 SET collation_connection  = latin1_swedish_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `RemoveScale` AFTER DELETE ON `Scales`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `RemoveScale`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `scheduler`
--

DROP TABLE IF EXISTS `scheduler`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `scheduler` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Name` varchar(255) NOT NULL DEFAULT '',
  `ScaleType` int(11) DEFAULT '0',
  `ScaleNumber` int(11) DEFAULT '0',
  `Task` smallint(6) DEFAULT '0',
  `Type` smallint(6) DEFAULT '0',
  `DatTim` datetime NOT NULL DEFAULT '1970-01-01 00:00:00',
  `LastExec` datetime NOT NULL DEFAULT '1970-01-01 00:00:00',
  `TaskParam` varchar(255) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `scheduler`
--

LOCK TABLES `scheduler` WRITE;
/*!40000 ALTER TABLE `scheduler` DISABLE KEYS */;
/*!40000 ALTER TABLE `scheduler` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `scripts`
--

DROP TABLE IF EXISTS `scripts`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `scripts` (
  `Code` int(11) NOT NULL,
  `Description` varchar(127) DEFAULT '',
  `Script` text,
  PRIMARY KEY (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `scripts`
--

LOCK TABLES `scripts` WRITE;
/*!40000 ALTER TABLE `scripts` DISABLE KEYS */;
/*!40000 ALTER TABLE `scripts` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertScripts` AFTER INSERT ON `Scripts`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckScripts`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateScripts` AFTER UPDATE ON `Scripts`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckScripts`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteScripts` AFTER DELETE ON `Scripts`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckScripts`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `sections`
--

DROP TABLE IF EXISTS `sections`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `sections` (
  `Code` decimal(3,0) NOT NULL,
  `Name` varchar(20) NOT NULL DEFAULT '',
  `Icon` varchar(127) DEFAULT NULL,
  `IsFrozen` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`Code`),
  KEY `Name` (`Name`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `sections`
--

LOCK TABLES `sections` WRITE;
/*!40000 ALTER TABLE `sections` DISABLE KEYS */;
/*!40000 ALTER TABLE `sections` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertSection` AFTER INSERT ON `Sections`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `CheckSections`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateSection` AFTER UPDATE ON `Sections`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `CheckSections`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteSection` AFTER DELETE ON `Sections`
FOR EACH ROW BEGIN
	UPDATE `LocalStatus` SET `CheckSections`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `selfctl`
--

DROP TABLE IF EXISTS `selfctl`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `selfctl` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Number` int(11) DEFAULT '0',
  `PrintDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `Relation` int(11) DEFAULT NULL,
  `Code` int(11) DEFAULT NULL,
  `SectionItem` int(11) DEFAULT '0',
  `Name` varchar(127) NOT NULL,
  `DateTime` datetime DEFAULT NULL,
  `Weight` decimal(5,3) DEFAULT NULL,
  `Units` int(11) DEFAULT NULL,
  `Status` smallint(6) DEFAULT '0',
  PRIMARY KEY (`Id`),
  KEY `Relation` (`Relation`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `selfctl`
--

LOCK TABLES `selfctl` WRITE;
/*!40000 ALTER TABLE `selfctl` DISABLE KEYS */;
/*!40000 ALTER TABLE `selfctl` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `selfctloper`
--

DROP TABLE IF EXISTS `selfctloper`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `selfctloper` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `OpType` char(1) DEFAULT '0',
  `Code` smallint(6) DEFAULT '0',
  `Name` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `OpType` (`OpType`,`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `selfctloper`
--

LOCK TABLES `selfctloper` WRITE;
/*!40000 ALTER TABLE `selfctloper` DISABLE KEYS */;
/*!40000 ALTER TABLE `selfctloper` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `selfctlrelations`
--

DROP TABLE IF EXISTS `selfctlrelations`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `selfctlrelations` (
  `IdRel` int(11) NOT NULL,
  `Source` smallint(6) DEFAULT '0',
  `Type` smallint(6) DEFAULT '0',
  `Corrective` smallint(6) DEFAULT '0',
  `Name` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`IdRel`),
  UNIQUE KEY `Source` (`Source`,`Type`,`Corrective`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `selfctlrelations`
--

LOCK TABLES `selfctlrelations` WRITE;
/*!40000 ALTER TABLE `selfctlrelations` DISABLE KEYS */;
/*!40000 ALTER TABLE `selfctlrelations` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `shoporder`
--

DROP TABLE IF EXISTS `shoporder`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `shoporder` (
  `Code` int(11) NOT NULL,
  `Name` varchar(127) NOT NULL,
  `SaleForm` smallint(6) DEFAULT '1',
  `Weight` decimal(12,3) DEFAULT '0.000',
  `Units` int(11) DEFAULT '0',
  `Observations` varchar(127) DEFAULT NULL,
  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  PRIMARY KEY (`Code`),
  CONSTRAINT `shoporder_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `shoporder`
--

LOCK TABLES `shoporder` WRITE;
/*!40000 ALTER TABLE `shoporder` DISABLE KEYS */;
/*!40000 ALTER TABLE `shoporder` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `sservicelabels`
--

DROP TABLE IF EXISTS `sservicelabels`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `sservicelabels` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Captured` char(1) DEFAULT '',
  `CapturedDateTime` datetime NOT NULL DEFAULT '1970-01-01 00:00:00',
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `lineType` smallint(6) DEFAULT NULL,
  `Item` int(11) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `Weight` decimal(7,3) DEFAULT NULL,
  `Tare` decimal(7,3) DEFAULT '0.000',
  `TTare` char(1) DEFAULT '',
  `Price` decimal(8,2) DEFAULT NULL,
  `lineDiscount` decimal(4,2) DEFAULT NULL,
  `Amount` decimal(10,2) DEFAULT NULL,
  `PacketExt` char(1) NOT NULL DEFAULT '-',
  `BovineCard` int(11) DEFAULT NULL,
  `EarTag` varchar(30) DEFAULT NULL,
  `BovineLot` varchar(20) DEFAULT NULL,
  `FreezingDate` date DEFAULT NULL,
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `DatTim` datetime DEFAULT NULL,
  `EAN13` varchar(13) NOT NULL DEFAULT '',
  `Code128` varchar(32) NOT NULL DEFAULT '',
  `Lot` varchar(20) NOT NULL DEFAULT '',
  `Ordinal` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `sservicelabels`
--

LOCK TABLES `sservicelabels` WRITE;
/*!40000 ALTER TABLE `sservicelabels` DISABLE KEYS */;
/*!40000 ALTER TABLE `sservicelabels` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `status`
--

DROP TABLE IF EXISTS `status`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `status` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Version` smallint(6) DEFAULT '14',
  `AppVersion` varchar(16) DEFAULT '',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `status`
--

LOCK TABLES `status` WRITE;
/*!40000 ALTER TABLE `status` DISABLE KEYS */;
INSERT INTO `status` VALUES (1,111,'2.20.3003');
/*!40000 ALTER TABLE `status` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `stock`
--

DROP TABLE IF EXISTS `stock`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `stock` (
  `Code` int(11) NOT NULL,
  `Name` varchar(127) NOT NULL,
  `SaleForm` smallint(6) DEFAULT '1',
  `Weight` decimal(12,3) DEFAULT '0.000',
  `Units` int(11) DEFAULT '0',
  `Inventory_Weight` decimal(12,3) DEFAULT '0.000',
  `Inventory_Units` int(11) DEFAULT '0',
  PRIMARY KEY (`Code`),
  CONSTRAINT `stock_ibfk_1` FOREIGN KEY (`Code`) REFERENCES `items` (`Code`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `stock`
--

LOCK TABLES `stock` WRITE;
/*!40000 ALTER TABLE `stock` DISABLE KEYS */;
/*!40000 ALTER TABLE `stock` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tarecodes`
--

DROP TABLE IF EXISTS `tarecodes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tarecodes` (
  `Code` decimal(2,0) NOT NULL,
  `Tare` decimal(7,3) DEFAULT '0.000',
  `Name` varchar(20) NOT NULL DEFAULT '',
  PRIMARY KEY (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tarecodes`
--

LOCK TABLES `tarecodes` WRITE;
/*!40000 ALTER TABLE `tarecodes` DISABLE KEYS */;
INSERT INTO `tarecodes` VALUES (1,1.000,'بيسس');
/*!40000 ALTER TABLE `tarecodes` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertCounterTareCodes` AFTER INSERT ON `TareCodes`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckTareCodes`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateCounterTareCodes` AFTER UPDATE ON `TareCodes`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckTareCodes`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteCounterTareCodes` AFTER DELETE ON `TareCodes`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckTareCodes`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `totaldaily`
--

DROP TABLE IF EXISTS `totaldaily`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `totaldaily` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` char(1) DEFAULT '',
  `NetStat` char(1) DEFAULT '',
  `Date` date NOT NULL,
  `Number` int(10) unsigned DEFAULT '0',
  `PrintChgCounter` int(10) unsigned DEFAULT '1',
  `Printing` tinyint(3) unsigned DEFAULT '0',
  `Printed` tinyint(3) unsigned DEFAULT '0',
  `Tickets` int(11) DEFAULT '1',
  `TotalLines` int(11) DEFAULT '0',
  `TotalAmount` decimal(12,2) DEFAULT '0.00',
  `TotalDiscount` decimal(12,2) DEFAULT '0.00',
  `TotalWeight` decimal(12,3) DEFAULT '0.000',
  `TotalUnits` int(11) DEFAULT '0',
  `NegativeLines` int(11) DEFAULT '0',
  `NegativeAmount` decimal(12,2) DEFAULT '0.00',
  `CanceledLines` int(11) DEFAULT '0',
  `CanceledAmount` decimal(12,2) DEFAULT '0.00',
  `Invoices` int(11) DEFAULT '0',
  `InvoiceAmount` decimal(12,2) DEFAULT '0.00',
  `Checksum` varchar(1024) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `dailycnt_idx` (`Date`,`Type`,`NetStat`)
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totaldaily`
--

LOCK TABLES `totaldaily` WRITE;
/*!40000 ALTER TABLE `totaldaily` DISABLE KEYS */;
INSERT INTO `totaldaily` VALUES (1,'T','N','2018-10-10',21,1,0,1,2,2,4.80,0.00,0.370,0,0,0.00,0,0.00,0,0.00,'43d0babaea3482041a25ea2868de775fca1c78a8e2dba77ae086ee976ad83831'),(2,'T','N','2018-12-12',22,1,0,1,3,4,30.08,0.00,0.115,3,0,0.00,0,0.00,0,0.00,'c25674e8c18a75f6bcf1e8eb41e6d9e60d25b743cbd15b28fb904d9c5f658b69'),(3,'T','N','2019-02-19',23,1,0,1,1,1,0.58,0.00,0.045,0,0,0.00,0,0.00,0,0.00,'c033ed68c9b84ff3dc1e3af76de2cbcb1d4517ab0d428480ef6d58122a1550a1'),(4,'T','N','2019-02-21',24,2,0,1,6,10,42.68,0.00,3.825,1,0,0.00,0,0.00,0,0.00,'855b9b69a4b21b4a16478a41c819524a25bb340748d1a013246fbc694bb2e1a9'),(7,'T','N','2019-02-22',25,2,0,0,6,9,16.89,0.00,3.095,2,0,0.00,0,0.00,1,10.47,'e526da49b077c481e860b3379a0f3a7bba0a580c7bd00134d1f9ddb9cf79d8dc'),(8,'T','N','2019-02-28',26,1,0,0,1,2,1.20,0.00,0.525,0,0,0.00,0,0.00,0,0.00,'15cbde10bd5cd934c933e55ed2ba39128a6527dc75f5a80fb581a40e0d2f2318');
/*!40000 ALTER TABLE `totaldaily` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertTotalDaily` AFTER INSERT ON `TotalDaily`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckTotalDaily`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateBTotalDaily` BEFORE UPDATE ON `TotalDaily`
FOR EACH ROW BEGIN
  IF((@DailyUnlock IS NULL) OR (@DailyUnlock!='451')) THEN 
    CALL `manipulate_protected_field`;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `totalpaytickets`
--

DROP TABLE IF EXISTS `totalpaytickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `totalpaytickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` char(1) DEFAULT 'T',
  `NetStat` char(1) DEFAULT '',
  `Date` date NOT NULL,
  `PayMethod` int(11) DEFAULT NULL,
  `Delivered` decimal(12,2) DEFAULT '0.00',
  `Checksum` varchar(1024) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `paytck_idx` (`Date`,`Type`,`NetStat`,`PayMethod`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalpaytickets`
--

LOCK TABLES `totalpaytickets` WRITE;
/*!40000 ALTER TABLE `totalpaytickets` DISABLE KEYS */;
/*!40000 ALTER TABLE `totalpaytickets` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateBTotalPayTickets` BEFORE UPDATE ON `TotalPayTickets`
FOR EACH ROW BEGIN
  IF((@DailyUnlock IS NULL) OR (@DailyUnlock!='451')) THEN 
    CALL `manipulate_protected_field`;  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `totalperpetual`
--

DROP TABLE IF EXISTS `totalperpetual`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `totalperpetual` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` char(1) DEFAULT '',
  `NetStat` char(1) DEFAULT '',
  `Date` date NOT NULL,
  `Tickets` int(11) DEFAULT '1',
  `TotalLines` int(11) DEFAULT '0',
  `TotalAmount` decimal(12,2) DEFAULT '0.00',
  `TotalDiscount` decimal(12,2) DEFAULT '0.00',
  `TotalWeight` decimal(12,3) DEFAULT '0.000',
  `TotalUnits` int(11) DEFAULT '0',
  `NegativeLines` int(11) DEFAULT '0',
  `NegativeAmount` decimal(12,2) DEFAULT '0.00',
  `CanceledLines` int(11) DEFAULT '0',
  `CanceledAmount` decimal(12,2) DEFAULT '0.00',
  `Invoices` int(11) DEFAULT '0',
  `InvoiceAmount` decimal(12,2) DEFAULT '0.00',
  `Checksum` varchar(1024) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `perpetual_idx` (`Type`,`NetStat`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalperpetual`
--

LOCK TABLES `totalperpetual` WRITE;
/*!40000 ALTER TABLE `totalperpetual` DISABLE KEYS */;
INSERT INTO `totalperpetual` VALUES (1,'T','N','2018-10-10',19,28,96.23,0.00,7.975,6,0,0.00,0,0.00,1,10.47,'8bf013e104db9830be0ae71afaf9c3b3e160cf418fb802de18e562647bc0cdd0');
/*!40000 ALTER TABLE `totalperpetual` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `totalscalefamilies`
--

DROP TABLE IF EXISTS `totalscalefamilies`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `totalscalefamilies` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Date` date NOT NULL,
  `Number` int(10) unsigned DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Family` int(11) DEFAULT '0',
  `PrintChgCounter` int(10) unsigned DEFAULT '1',
  `Printing` tinyint(3) unsigned DEFAULT '0',
  `Printed` tinyint(3) unsigned DEFAULT '0',
  `TotalWeight` decimal(12,3) DEFAULT '0.000',
  `TotalUnits` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `total_idx` (`Date`,`Scale`,`Family`)
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalscalefamilies`
--

LOCK TABLES `totalscalefamilies` WRITE;
/*!40000 ALTER TABLE `totalscalefamilies` DISABLE KEYS */;
INSERT INTO `totalscalefamilies` VALUES (1,'2016-05-01',1,1,101,1,0,0,0.000,1),(2,'2016-08-01',2,1,101,1,0,0,0.000,1),(3,'2016-10-01',3,1,501,1,0,0,0.000,1),(4,'2016-10-01',3,1,101,1,0,0,3.025,0),(5,'2016-10-01',4,2,101,1,0,0,0.220,0),(6,'2018-10-01',5,1,101,1,0,0,0.370,0),(7,'2018-12-01',6,1,0,1,0,0,0.115,0),(8,'2018-12-01',6,1,501,1,0,0,0.000,1),(9,'2018-12-01',6,1,101,1,0,0,0.000,2),(10,'2019-02-01',7,1,501,1,0,0,1.540,1),(11,'2019-02-01',7,1,201,1,0,0,1.700,2),(12,'2019-02-01',7,1,502,1,0,0,2.550,0),(13,'2019-02-01',7,1,105,1,0,0,0.425,0),(16,'2019-02-01',7,1,202,1,0,0,0.850,0),(17,'2019-02-01',7,1,203,1,0,0,0.425,0);
/*!40000 ALTER TABLE `totalscalefamilies` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertTotalScaleFamilies` AFTER INSERT ON `TotalScaleFamilies`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckTotalScaleFamilies`=TRUE;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `totalvat`
--

DROP TABLE IF EXISTS `totalvat`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `totalvat` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Type` char(1) DEFAULT 'T',
  `NetStat` char(1) DEFAULT '',
  `Date` date NOT NULL,
  `Code` decimal(1,0) unsigned NOT NULL,
  `Percentage` decimal(4,2) DEFAULT NULL,
  `Amount` decimal(12,2) DEFAULT '0.00',
  `Base` decimal(12,2) DEFAULT '0.00',
  `Tax` decimal(12,2) DEFAULT '0.00',
  `Checksum` varchar(1024) NOT NULL DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `totalvat_idx` (`Date`,`Type`,`NetStat`,`Code`)
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalvat`
--

LOCK TABLES `totalvat` WRITE;
/*!40000 ALTER TABLE `totalvat` DISABLE KEYS */;
INSERT INTO `totalvat` VALUES (1,'T','N','2018-10-10',0,0.00,4.80,4.80,0.00,'1da6802e509faab71cda24ff33af2406a19eae1038f6dbd5240ca98ac4b4f056'),(2,'T','N','2018-12-12',0,0.00,30.08,30.08,0.00,'2b1e9aafe73388cd35a016d544858a307d982b1775ddf627c43009700ce0a48a'),(3,'T','N','2019-02-19',0,0.00,0.58,0.58,0.00,'466f56a05d55c81bec11d4c171caca9d3da35d94d5136a39c8458982707eaac4'),(4,'T','N','2019-02-21',0,0.00,42.68,42.68,0.00,'e56f196c49e93d15ee317f889edefe3b42666c9f9ae5dda847dd19da3bca5042'),(7,'T','N','2019-02-22',0,0.00,16.89,16.89,0.00,'736da152a03883fcf8d9160f8112f5e096dadec1b39bd33fa2d309b164385c46'),(8,'T','N','2019-02-28',0,0.00,1.20,1.20,0.00,'285dd7e3f77996f135cd25f60b9669ab5e08acd6b1063eeee88714a338ee9c48');
/*!40000 ALTER TABLE `totalvat` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateBTotalVAT` BEFORE UPDATE ON `TotalVAT`
FOR EACH ROW BEGIN
  IF((@DailyUnlock IS NULL) OR (@DailyUnlock!='451')) THEN 
    CALL `manipulate_protected_field`;
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `trademark`
--

DROP TABLE IF EXISTS `trademark`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `trademark` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `CO_Name` varchar(255) DEFAULT '',
  `Description` varchar(255) DEFAULT '',
  `Business` varchar(255) DEFAULT '',
  `Slogan` varchar(255) DEFAULT '',
  `Headquarter` varchar(255) DEFAULT '',
  `CO_WEB` varchar(255) DEFAULT '',
  `Shop_Name` varchar(255) DEFAULT '',
  `Address` varchar(255) DEFAULT '',
  `City` varchar(255) DEFAULT '',
  `Postalcode` varchar(80) DEFAULT '',
  `Country` varchar(80) DEFAULT '',
  `Phone` varchar(80) DEFAULT '',
  `Fax` varchar(80) DEFAULT '',
  `MobilePhone` varchar(80) DEFAULT '',
  `eMail` varchar(80) DEFAULT '',
  `TIN` varchar(80) DEFAULT '',
  `Shop_WEB` varchar(255) DEFAULT '',
  `Text1` varchar(255) DEFAULT '',
  `Text2` varchar(255) DEFAULT '',
  `Text3` varchar(255) DEFAULT '',
  `Text4` varchar(255) DEFAULT '',
  `Text5` varchar(255) DEFAULT '',
  `Text6` varchar(255) DEFAULT '',
  `Text7` varchar(255) DEFAULT '',
  `Text8` varchar(255) DEFAULT '',
  `Text9` varchar(255) DEFAULT '',
  `Text10` varchar(255) DEFAULT '',
  `Icon` varchar(127) DEFAULT '',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `trademark`
--

LOCK TABLES `trademark` WRITE;
/*!40000 ALTER TABLE `trademark` DISABLE KEYS */;
INSERT INTO `trademark` VALUES (1,'','SUPERMARCADOS MASYMAS','','','','www.masymas.com','**** MASYMAS ****','SECCION','CARNICERIA','03750','ESPAÑA','965760450-630308329','','','fornes@fornes.net','CIF: A03140456','www.masymas.com/es','CRTA.VALENCIA-ALICANTE,KM.191.300','03750 PEDREGUER','Tel: 965760450','','','','','','','','');
/*!40000 ALTER TABLE `trademark` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`localhost`*/ /*!50003 TRIGGER `TrademarkInsertedB` BEFORE INSERT ON `Trademark`
FOR EACH ROW BEGIN
  SET NEW.`Id`=1;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `turn`
--

DROP TABLE IF EXISTS `turn`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `turn` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Name` varchar(32) NOT NULL DEFAULT '',
  `IP` int(10) unsigned NOT NULL DEFAULT '0',
  `Enabled` tinyint(1) DEFAULT '0',
  `User` varchar(32) NOT NULL DEFAULT '',
  `Password` varchar(32) NOT NULL DEFAULT '',
  `Status` char(1) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `IP` (`IP`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `turn`
--

LOCK TABLES `turn` WRITE;
/*!40000 ALTER TABLE `turn` DISABLE KEYS */;
/*!40000 ALTER TABLE `turn` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `TurnInserted` AFTER INSERT ON `Turn`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckTurn` = TRUE;
  SET @trx := (SELECT trx_new("TURN_RELOAD", "NETWORK"));
  SET @ret := (SELECT trx_send(@trx));
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `TurnUpdated` AFTER UPDATE ON `Turn`
FOR EACH ROW BEGIN
  IF ((NEW.`Enabled`!=OLD.`Enabled`) OR (NEW.`Status`!=OLD.`Status`)) THEN
    UPDATE `LocalStatus` SET `CheckTurn` = TRUE;
  END IF;
  IF ((NEW.`Enabled`!=OLD.`Enabled`) OR (NEW.`IP`!=OLD.`IP`) OR (NEW.`Name`!=OLD.`Name`)) THEN
    SET @trx := (SELECT trx_new("TURN_RELOAD", "NETWORK"));
    SET @ret := (SELECT trx_send(@trx));
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `TurnDeleted` AFTER DELETE ON `Turn`
FOR EACH ROW BEGIN
  UPDATE `LocalStatus` SET `CheckTurn` = TRUE;
  SET @trx := (SELECT trx_new("TURN_RELOAD", "NETWORK"));
  SET @ret := (SELECT trx_send(@trx));
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `turnswaiting`
--

DROP TABLE IF EXISTS `turnswaiting`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `turnswaiting` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Counter` decimal(3,0) DEFAULT '0',
  `DayWeek` decimal(1,0) DEFAULT '0',
  `TimeLimit` int(11) NOT NULL DEFAULT '0',
  `TurnsLimit` int(11) NOT NULL DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `idx_twt` (`Counter`,`DayWeek`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `turnswaiting`
--

LOCK TABLES `turnswaiting` WRITE;
/*!40000 ALTER TABLE `turnswaiting` DISABLE KEYS */;
/*!40000 ALTER TABLE `turnswaiting` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `users`
--

DROP TABLE IF EXISTS `users`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `users` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `User` varchar(16) NOT NULL DEFAULT '',
  `Password` varchar(31) NOT NULL,
  `Name` varchar(127) NOT NULL,
  `LevelLab` tinyint(1) DEFAULT '0',
  `LevelAdmin` tinyint(1) DEFAULT '0',
  `LevelConfig` tinyint(1) DEFAULT '0',
  `LevelProg` tinyint(1) DEFAULT '0',
  `LevelVendor` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `User` (`User`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `users`
--

LOCK TABLES `users` WRITE;
/*!40000 ALTER TABLE `users` DISABLE KEYS */;
INSERT INTO `users` VALUES (1,'admin','YWRtaW4=','Administrator',0,1,0,0,0),(2,'sales','c2FsZXM=','Sales user',0,0,0,0,1),(3,'packet','cGFja2V0','Packet user',0,0,0,0,1),(4,'sservice','c3NlcnZpY2U=','Self-Service user',0,0,0,0,0);
/*!40000 ALTER TABLE `users` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `vats`
--

DROP TABLE IF EXISTS `vats`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `vats` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Code` decimal(1,0) unsigned NOT NULL,
  `Name` varchar(16) NOT NULL,
  `Percent` decimal(4,2) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Code` (`Code`)
) ENGINE=InnoDB AUTO_INCREMENT=23 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vats`
--

LOCK TABLES `vats` WRITE;
/*!40000 ALTER TABLE `vats` DISABLE KEYS */;
INSERT INTO `vats` VALUES (3,1,'IVA 10 %',10.00),(4,2,'IVA 18 %',18.00),(22,0,'Sin IVA',0.00);
/*!40000 ALTER TABLE `vats` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `vendorkeys`
--

DROP TABLE IF EXISTS `vendorkeys`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `vendorkeys` (
  `Counter` int(11) DEFAULT NULL,
  `VendorKey` int(11) DEFAULT NULL,
  `Vendor` int(11) NOT NULL,
  PRIMARY KEY (`Vendor`),
  UNIQUE KEY `Counter` (`Counter`,`VendorKey`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vendorkeys`
--

LOCK TABLES `vendorkeys` WRITE;
/*!40000 ALTER TABLE `vendorkeys` DISABLE KEYS */;
INSERT INTO `vendorkeys` VALUES (2,1,52001),(2,2,52002),(2,3,52003),(2,4,52004),(2,5,52005),(2,6,52006),(2,7,52007),(2,8,52008),(2,9,52009),(2,10,52000),(3,1,50001),(3,2,50002),(3,3,50003),(3,4,50004),(3,5,50005),(3,6,50006),(3,7,50007),(3,8,50008),(3,9,50009),(3,10,50000),(4,1,54001),(4,2,54002),(4,3,54003),(4,4,54004),(4,5,54005),(4,6,54006),(4,7,54007),(4,8,54008),(4,9,54009),(4,10,54000),(5,2,58002),(5,3,58003),(5,4,58004),(5,5,58005),(5,6,58006),(5,7,58007),(5,8,58008),(5,9,58009),(5,10,58000),(6,1,56001),(6,2,56002),(6,3,56003),(6,4,56004),(6,5,56005),(6,6,56006),(6,7,56007),(6,8,56008),(6,9,56009),(6,10,56000),(8,1,59001),(8,2,59002),(8,3,59003),(8,4,59004),(8,5,59005),(8,6,59006),(8,7,59007),(8,8,59008),(8,9,59009);
/*!40000 ALTER TABLE `vendorkeys` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertVendorKey` AFTER INSERT ON `VendorKeys`
FOR EACH ROW BEGIN
  SET @disable_localregister = 1;
  REPLACE INTO `LocalVendorKeys` (`Counter`, `VendorKey`, `Vendor`) VALUES( NEW.`Counter`, NEW.`VendorKey`, NEW.`Vendor`);
  IF @disable_netregister IS NULL THEN
    INSERT INTO `VendorRegister` (`Vendor`,`RegisterDateTime`) VALUES (NEW.`Vendor`,NOW());
  END IF;
  SET @disable_localregister = NULL;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteVendorKey` AFTER DELETE ON `VendorKeys`
FOR EACH ROW BEGIN
  DELETE FROM `LocalVendorKeys` WHERE `Vendor`=OLD.`Vendor`;
  UPDATE `VendorRegister` SET `ReleaseDateTime`=NOW() WHERE `Vendor`=OLD.`Vendor` AND `ReleaseDateTime`<='1970-01-01 00:00:00';
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `vendorregister`
--

DROP TABLE IF EXISTS `vendorregister`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `vendorregister` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Vendor` int(11) DEFAULT NULL,
  `RegisterDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `RegisterSend` smallint(6) DEFAULT '0',
  `ReleaseDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReleaseSend` smallint(6) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=178 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vendorregister`
--

LOCK TABLES `vendorregister` WRITE;
/*!40000 ALTER TABLE `vendorregister` DISABLE KEYS */;
/*!40000 ALTER TABLE `vendorregister` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `vendors`
--

DROP TABLE IF EXISTS `vendors`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `vendors` (
  `Code` int(11) NOT NULL,
  `Name` varchar(64) NOT NULL,
  `Password` varchar(31) NOT NULL DEFAULT '',
  `Icon` varchar(127) NOT NULL DEFAULT '',
  `TurnId` int(10) unsigned NOT NULL DEFAULT '0',
  PRIMARY KEY (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vendors`
--

LOCK TABLES `vendors` WRITE;
/*!40000 ALTER TABLE `vendors` DISABLE KEYS */;
INSERT INTO `vendors` VALUES (50000,'CHARCU 10                    .','','',0),(50001,'CHARCU 1                     .','','',0),(50002,'CHARCU 2                     .','','',0),(50003,'CHARCU 3                     .','','',0),(50004,'CHARCU 4                     .','','',0),(50005,'CHARCU 5                     .','','',0),(50006,'CHARCU 6                     .','','',0),(50007,'CHARCU 7                     .','','',0),(50008,'CHARCU 8                     .','','',0),(50009,'CHARCU 9                     .','','',0),(52000,'CARNE 10                     .','','',0),(52001,'CARNE 1                      .','','',0),(52002,'CARNE 2                      .','','',0),(52003,'CARNE 3                      .','','',0),(52004,'CARNE 4                      .','','',0),(52005,'CARNE 5                      .','','',0),(52006,'CARNE 6                      .','','',0),(52007,'CARNE 7                      .','','',0),(52008,'CARNE 8                      .','','',0),(52009,'CARNE 9                      .','','',0),(54000,'FYV 10                       .','','',0),(54001,'FYV 1                        .','','',0),(54002,'FYV 2                        .','','',0),(54003,'FYV 3                        .','','',0),(54004,'FYV 4                        .','','',0),(54005,'FYV 5                        .','','',0),(54006,'FYV 6                        .','','',0),(54007,'FYV 7                        .','','',0),(54008,'FYV 8                        .','','',0),(54009,'FYV 9                        .','','',0),(56000,'PAN 10                       .','','',0),(56001,'PAN 1                        .','','',0),(56002,'PAN 2                        .','','',0),(56003,'PAN 3                        .','','',0),(56004,'PAN 4                        .','','',0),(56005,'PAN 5                        .','','',0),(56006,'PAN 6                        .','','',0),(56007,'PAN 7                        .','','',0),(56008,'PAN 8                        .','','',0),(56009,'PAN 9                        .','','',0),(58000,'PESCA 10                     .','','',0),(58002,'PESCA 2                      .','','',0),(58003,'PESCA 3                      .','','',0),(58004,'PESCA 4                      .','','',0),(58005,'PESCA 5                      .','','',0),(58006,'PESCA 6                      .','','',0),(58007,'PESCA 7                      .','','',0),(58008,'PESCA 8                      .','','',0),(58009,'PESCA 9                      .','','',0),(59001,'COCINA 1                     .','','',0),(59002,'COCINA 2                     .','','',0),(59003,'COCINA 3                     .','','',0),(59004,'COCINA 4                     .','','',0),(59005,'COCINA 5                     .','','',0),(59006,'COCINA 6                     .','','',0),(59007,'COCINA 7                     .','','',0),(59008,'COCINA 8                     .','','',0),(59009,'COCINA 9                     .','','',0);
/*!40000 ALTER TABLE `vendors` ENABLE KEYS */;
UNLOCK TABLES;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `InsertVendor` AFTER INSERT ON `Vendors`
FOR EACH ROW BEGIN
  IF (SELECT `Counter` FROM `VendorKeys` WHERE `VendorKeys`.`Vendor`=NEW.`Code`AND `VendorKeys`.`Counter`=GLOBAL_GET("current_counter")) IS NOT NULL THEN
	REPLACE INTO `LocalVendorChanges` VALUES (NEW.`Code`*10, NEW.`Code`,'0');
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `UpdateVendor` AFTER UPDATE ON `Vendors`
FOR EACH ROW BEGIN
  IF (OLD.`Code`!=NEW.`Code`) THEN
      REPLACE INTO `LocalVendorChanges` VALUES ((OLD.`Code`*10)+1, OLD.`Code`,'1');
  END IF;
  REPLACE INTO `LocalVendorChanges` VALUES (NEW.`Code`*10, NEW.`Code`,'0');
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `DeleteVendor` AFTER DELETE ON `Vendors`
FOR EACH ROW BEGIN
  IF (SELECT `Counter` FROM `VendorKeys` WHERE `VendorKeys`.`Vendor`=OLD.`Code`AND `VendorKeys`.`Counter`=GLOBAL_GET("current_counter")) IS NOT NULL THEN
    REPLACE INTO `LocalVendorChanges` VALUES ((OLD.`Code`*10)+1, OLD.`Code`,'1');
  END IF;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

--
-- Table structure for table `vitaminnames`
--

DROP TABLE IF EXISTS `vitaminnames`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `vitaminnames` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `locale` char(2) NOT NULL DEFAULT '',
  `type` char(1) DEFAULT NULL,
  `shortname` char(3) NOT NULL DEFAULT '',
  `name` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `locale` (`locale`,`type`,`shortname`)
) ENGINE=InnoDB AUTO_INCREMENT=82 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vitaminnames`
--

LOCK TABLES `vitaminnames` WRITE;
/*!40000 ALTER TABLE `vitaminnames` DISABLE KEYS */;
INSERT INTO `vitaminnames` VALUES (1,'es','V','A','Vitamina A'),(2,'es','V','D','Vitamina D'),(3,'es','V','E','Vitamina E'),(4,'es','V','K','Vitamina K'),(5,'es','V','C','Vitamina C'),(6,'es','V','B1','Tiamina'),(7,'es','V','B2','Riboflavina'),(8,'es','V','B3','Niacina'),(9,'es','V','B6','Vitamina B6'),(10,'es','V','B9','Ácido fólico'),(11,'es','V','B12','Vitamina B12'),(12,'es','V','B8','Biotina'),(13,'es','V','B5','Ácido pantoténico'),(14,'es','M','K','Potasio'),(15,'es','M','Cl','Cloro'),(16,'es','M','Ca','Calcio'),(17,'es','M','P','Fósforo'),(18,'es','M','Mg','Magnesio'),(19,'es','M','Fe','Hierro'),(20,'es','M','Zn','Zinc'),(21,'es','M','Cu','Cobre'),(22,'es','M','Mn','Manganeso'),(23,'es','M','F','Flúor'),(24,'es','M','Se','Selenio'),(25,'es','M','Cr','Cromo'),(26,'es','M','Mo','Molibdeno'),(27,'es','M','I','Yodo'),(28,'en','V','A','Vitamin A'),(29,'en','V','D','Vitamin D'),(30,'en','V','E','Vitamin E'),(31,'en','V','K','Vitamin K'),(32,'en','V','C','Vitamin C'),(33,'en','V','B1','Thiamin'),(34,'en','V','B2','Riboflavin'),(35,'en','V','B3','Niacin'),(36,'en','V','B6','Vitamin B6'),(37,'en','V','B9','Folic acid'),(38,'en','V','B12','Vitamin B12'),(39,'en','V','B8','Biotin'),(40,'en','V','B5','Pantothenic acid'),(41,'en','M','K','Potassium'),(42,'en','M','Cl','Chloride'),(43,'en','M','Ca','Calcium'),(44,'en','M','P','Phosphorus'),(45,'en','M','Mg','Magnesium'),(46,'en','M','Fe','Iron'),(47,'en','M','Zn','Zinc'),(48,'en','M','Cu','Copper'),(49,'en','M','Mn','Manganese'),(50,'en','M','F','Fluoride'),(51,'en','M','Se','Selenium'),(52,'en','M','Cr','Chromium'),(53,'en','M','Mo','Molybdenum'),(54,'en','M','I','Iodine'),(55,'ca','V','A','Vitamina A'),(56,'ca','V','D','Vitamina D'),(57,'ca','V','E','Vitamina E'),(58,'ca','V','K','Vitamina K'),(59,'ca','V','C','Vitamina C'),(60,'ca','V','B1','Tiamina'),(61,'ca','V','B2','Riboflavina'),(62,'ca','V','B3','Niacina'),(63,'ca','V','B6','Vitamina B6'),(64,'ca','V','B9','Àcid fòlic'),(65,'ca','V','B12','Vitamina B12'),(66,'ca','V','B8','Biotina'),(67,'ca','V','B5','Àcid pantotènic'),(68,'ca','M','K','Potassi'),(69,'ca','M','Cl','Clorur'),(70,'ca','M','Ca','Calci'),(71,'ca','M','P','Fòsfor'),(72,'ca','M','Mg','Magnesi'),(73,'ca','M','Fe','Ferro'),(74,'ca','M','Zn','Zinc'),(75,'ca','M','Cu','Coure'),(76,'ca','M','Mn','Manganès'),(77,'ca','M','F','Fluorur'),(78,'ca','M','Se','Seleni'),(79,'ca','M','Cr','Crom'),(80,'ca','M','Mo','Molibdè'),(81,'ca','M','I','Iode');
/*!40000 ALTER TABLE `vitaminnames` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `vitamins`
--

DROP TABLE IF EXISTS `vitamins`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `vitamins` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `card` int(11) NOT NULL,
  `identifier` int(11) NOT NULL,
  `type` char(1) DEFAULT NULL,
  `shortname` char(3) NOT NULL DEFAULT '',
  `name` varchar(50) DEFAULT NULL,
  `units` varchar(2) DEFAULT NULL,
  `daily_ref` float DEFAULT NULL,
  `per100` float DEFAULT '0',
  `perportion` float DEFAULT '0',
  PRIMARY KEY (`id`),
  UNIQUE KEY `vtm` (`card`,`identifier`),
  CONSTRAINT `vitamins_ibfk_1` FOREIGN KEY (`card`) REFERENCES `nutrition` (`card`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vitamins`
--

LOCK TABLES `vitamins` WRITE;
/*!40000 ALTER TABLE `vitamins` DISABLE KEYS */;
/*!40000 ALTER TABLE `vitamins` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `xmltexts`
--

DROP TABLE IF EXISTS `xmltexts`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `xmltexts` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `XmlType` char(1) DEFAULT NULL,
  `GroupName` varchar(32) NOT NULL DEFAULT '',
  `Name` varchar(32) NOT NULL DEFAULT '',
  `XmlText` text NOT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `XmlType` (`XmlType`,`GroupName`,`Name`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `xmltexts`
--

LOCK TABLES `xmltexts` WRITE;
/*!40000 ALTER TABLE `xmltexts` DISABLE KEYS */;
/*!40000 ALTER TABLE `xmltexts` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Dumping routines for database 'pcscale'
--
/*!50003 DROP PROCEDURE IF EXISTS `dbTrigger` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
CREATE DEFINER=`pcscale`@`%` PROCEDURE `dbTrigger`(IN `tbname` VARCHAR (40), `function` VARCHAR (10), `masterkey` VARCHAR (40))
BEGIN
    SET @trx := (SELECT trx_new("DBTRIGGER", "CORE"));
    SET @ret := (SELECT trx_set_param(@trx, "TABLE", `tbname`));
    SET @ret := (SELECT trx_set_param(@trx, "FUNCTION", `function`));
    SET @ret := (SELECT trx_set_param(@trx, "KEY", `masterkey`));
    SET @ret := (SELECT trx_send(@trx));
END ;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2019-02-28 19:37:41
