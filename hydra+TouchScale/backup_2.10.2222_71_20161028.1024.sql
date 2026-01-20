-- MySQL dump 10.14  Distrib 5.5.46-MariaDB, for debian-linux-gnu (i686)
--
-- Host: localhost    Database: pcscale
-- ------------------------------------------------------
-- Server version	5.5.46-MariaDB-1~precise-log

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
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `barcodeean13`
--

LOCK TABLES `barcodeean13` WRITE;
/*!40000 ALTER TABLE `barcodeean13` DISABLE KEYS */;
INSERT INTO `barcodeean13` VALUES (1,'G',0,'','20VVViiiiiii','10VVVIIIIIII','2YTTTTTiiiii','2CCCCCCiiiii','2CCCCCCiiiii','','20VVViiiiiii','10VVVIIIIIII','2YTTTTTiiiii','2CCCCCCiiiii','2CCCCCCiiiii','2CCCCCCIIIII');
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
/*!40000 ALTER TABLE `bovine` ENABLE KEYS */;
UNLOCK TABLES;

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
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `counterfamilies`
--

LOCK TABLES `counterfamilies` WRITE;
/*!40000 ALTER TABLE `counterfamilies` DISABLE KEYS */;
INSERT INTO `counterfamilies` VALUES (1,1,1,1),(2,1,2,2),(3,1,3,3),(4,1,4,4),(5,1,5,5),(6,1,6,6),(7,1,7,9),(8,1,8,462),(9,1,9,463),(10,1,10,464),(11,1,11,465),(12,1,12,466),(13,1,13,467),(14,1,14,468),(15,1,15,469);
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
  `Turn` int(10) unsigned DEFAULT '1',
  `Description` varchar(127) DEFAULT '',
  PRIMARY KEY (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `counters`
--

LOCK TABLES `counters` WRITE;
/*!40000 ALTER TABLE `counters` DISABLE KEYS */;
INSERT INTO `counters` VALUES (1,'',NULL,1,'');
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
  PRIMARY KEY (`Id`),
  UNIQUE KEY `TIN` (`TIN`),
  KEY `Name` (`Name`(19))
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `customers`
--


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
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `erasetickets`
--

LOCK TABLES `erasetickets` WRITE;
/*!40000 ALTER TABLE `erasetickets` DISABLE KEYS */;
INSERT INTO `erasetickets` VALUES (1,'T','2016-10-28 10:16:29'),(2,'P','2016-10-28 10:07:44'),(3,'S','2016-10-28 10:07:40');
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
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
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
INSERT INTO `families` VALUES (1,'CITRICOS','13725_1.png','',0,''),(2,'FRUTAS GRANDES','fruits2 625x348.png','',0,''),(3,'FRUTAS PEQUEÑAS','fruits 625x348.png','',0,''),(4,'FRUTAS TROPICALES','10009.jpg','',0,''),(5,'FRUTOS SECOS','13779.png','',0,''),(6,'HORTALIZAS','13806.png','',0,''),(9,'VERDURAS','13832.png','',0,''),(462,'REFRESCOS','fontvella.png','',0,''),(463,'AROMATICOS','IMG_2397.JPG','',0,''),(464,'ENSALADAS','V_062.png','',0,''),(465,'SETAS','Setas.jpg','',0,''),(466,'ELABORADOS CASA CESPEDES','MACEDONIA.JPG','',0,''),(467,'HUEVOS CREMAS CONSERVAS','1379430312200.jpg','',0,''),(468,'PRECOCINADOS','Menestra.jpg','',0,''),(469,'PRODUCTOS PAIS VASCO','','',0,'');
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
  `SectionItem` int(11) DEFAULT '0',
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
  `StartDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `TimesPrinted` tinyint(3) unsigned NOT NULL DEFAULT '0',
  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadFrom` int(11) DEFAULT '0',
  `ReloadScale` int(11) NOT NULL DEFAULT '0',
  `ReloadDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadTicketNumber` int(11) NOT NULL DEFAULT '0',
  `ReloadtckType` char(1) DEFAULT '',
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
  `Invoice` varchar(20) NOT NULL DEFAULT '',
  `InvoiceDate` date DEFAULT '0000-00-00',
  `tckDiscount` decimal(4,2) DEFAULT '0.00',
  `ParkedNumber` int(11) DEFAULT '0',
  `NeedCopy` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`Id`),
  KEY `Vendor` (`Vendor`),
  KEY `Invoice` (`Invoice`),
  KEY `datim_idx` (`DatTim`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `htickets`
--

LOCK TABLES `htickets` WRITE;
/*!40000 ALTER TABLE `htickets` DISABLE KEYS */;
INSERT INTO `htickets` VALUES (1,0,'N','*','0000-00-00 00:00:00','T','C',1,1,1,105,0,2,2,13.58,10.19,4.000,0,0,0.00,0,0.00,'0000000000000','2016-10-28 10:20:56',0,'2016-10-28 10:21:13',0,0,'0000-00-00 00:00:00',0,'','',1,'160010100107','2016-10-28',25.00,0,0),(2,0,'N','*','0000-00-00 00:00:00','T','C',2,1,1,105,0,2,2,3.50,3.50,2.000,1,0,0.00,0,0.00,'0000000000000','2016-10-28 10:21:27',0,'2016-10-28 10:21:44',0,0,'0000-00-00 00:00:00',0,'','',1,'160010100108','2016-10-28',0.00,0,0),(3,0,'N','*','0000-00-00 00:00:00','T','C',3,1,1,105,0,1,1,10.00,9.00,2.000,0,0,0.00,0,0.00,'0000000000000','2016-10-28 10:21:58',0,'2016-10-28 10:22:06',0,0,'0000-00-00 00:00:00',0,'','',1,'160010100109','2016-10-28',10.00,0,0);
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
  IF(NEW.`NetStat`='N') AND (OLD.`tckDiscount`!=NEW.`tckDiscount`) THEN
    SET @frac_digits := IFNULL(GLOBAL_GET("frac_digits"),2);
    SET NEW.`NetAmount`=ROUND(NEW.`GrossAmount`*((100-NEW.`tckDiscount`)/100), @frac_digits);
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
  SET @frac_digits := IFNULL(GLOBAL_GET("frac_digits"),2);

  IF (NEW.`NetStat`='N') AND (NEW.`Counter`= GLOBAL_GET("current_counter")) THEN
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
    IF (NEW.`Status`='C') THEN

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

      UPDATE `Stock`,
        (SELECT `Item`, SUM(IF(`Saleform`=1,`Weight`,0)) AS `lwght`, SUM(IF(`Saleform`=0,`Weight`,0)) AS `lunits` FROM `LTickets` WHERE (`IdHTicket`=NEW.`Id` AND `LTickets`.`lineType`=1) GROUP BY `Item`) AS `VT_L`
        SET `Weight`=IF(`Weight`>`VT_L`.`lwght`, `Weight`-`VT_L`.`lwght`, 0), `Stock`.`Units`=IF(`Stock`.`Units`>`VT_L`.`lunits`, `Stock`.`Units`-`VT_L`.`lunits`, 0) WHERE `Code` = `VT_L`.`Item`;

      IF (NEW.`tckType`!='U') THEN


        INSERT INTO `TotalVAT` (`Date`, `NetStat`, `Code`, `Percentage`, `Amount`, `Base`, `Tax`)
          SELECT DATE(NEW.`DatTim`), NEW.`NetStat`, `V_VAT`.`VATCode`, `V_VAT`.`VATPercent`, `V_VAT`.`Amount`,
          ROUND((`V_VAT`.`Amount`*100)/(100+`VATPercent`),@frac_digits) AS `Base`, `V_VAT`.`Amount` - ROUND((`V_VAT`.`Amount`*100)/(100+`VATPercent`),@frac_digits) AS Tax
          FROM (SELECT `VATCode`, `VATPercent`, ROUND(SUM(`LTickets`.`Amount`)*((100-`LTickets`.`TicketDiscount`)/100), @frac_digits) AS `Amount`
            FROM `LTickets` WHERE `LTickets`.`IdHTicket`=NEW.`Id` AND `LTickets`.`lineType`=1 GROUP BY `LTickets`.`VATCode`) AS `V_VAT`
            ON DUPLICATE KEY UPDATE `TotalVAT`.`Amount`=`TotalVAT`.`Amount`+VALUES(`Amount`), `TotalVAT`.`Base`=`TotalVAT`.`Base`+VALUES(`Base`),
            `TotalVAT`.`Tax`=`TotalVAT`.`Tax`+VALUES(`Tax`);

        INSERT INTO `TotalPayTickets` (`Date`, `NetStat`, `PayMethod`, `Delivered`)
          (SELECT DATE(NEW.`DatTim`), NEW.`NetStat`,`PayMethod`, `Delivered` FROM `PayTickets` WHERE `PayTickets`.`IdHTicket`=NEW.`Id`)
          ON DUPLICATE KEY UPDATE `TotalPayTickets`.`Delivered`=`TotalPayTickets`.`Delivered`+VALUES(`Delivered`);

        IF (OLD.`SimplifiedInvoice`!=0) THEN
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

    IF(OLD.`Status`='C') THEN


      UPDATE `TotalDaily` SET `PrintChgCounter`=`PrintChgCounter`+`Printed`, `Printed`=0,
        `Tickets`=`Tickets`-1, `TotalLines`=`TotalLines`-OLD.`NumLines`,`TotalAmount`=`TotalAmount`-OLD.`NetAmount`,
        `TotalDiscount`=`TotalDiscount`-(OLD.`GrossAmount`-OLD.`NetAmount`), `TotalWeight`=`TotalWeight`-OLD.`TotalWeight`, `TotalUnits`=`TotalUnits`-OLD.`TotalUnits`,
        `NegativeLines`=`NegativeLines`-OLD.`NegativeLines`, `NegativeAmount`=`NegativeAmount`-OLD.`NegativeAmount`,
        `CanceledLines`=`CanceledLines`-OLD.`CanceledLines`, `CanceledAmount`=`CanceledAmount`-OLD.`CanceledAmount`,
        `Invoices`=`Invoices`-IF(OLD.`Invoice`!='',1,0), `InvoiceAmount`=`InvoiceAmount`-IF(OLD.`Invoice`!='',OLD.`NetAmount`,0)
        WHERE `Date`=DATE(OLD.`DatTim`) AND `Type`=OLD.`tckType` AND `NetStat`=OLD.`NetStat`;

      UPDATE `Stock`,
        (SELECT `Item`, SUM(IF(`Saleform`=1,`Weight`,0)) AS `lwght`, SUM(IF(`Saleform`=0,`Weight`,0)) AS `lunits` FROM `LTickets` WHERE (`IdHTicket`=NEW.`Id`  AND `LTickets`.`lineType`=1) GROUP BY `Item`) AS `VT_L`
        SET `Weight`=`Stock`.`Weight`+`VT_L`.`lwght`, `Stock`.`Units`=`Stock`.`Units`+`VT_L`.`lunits` WHERE `Code` = `VT_L`.`Item`;

      IF (OLD.`tckType`!='U') THEN


        UPDATE `TotalVAT`,
          (SELECT `VATCode`, `VATPercent`, ROUND(SUM(`LTickets`.`Amount`)*((100-`LTickets`.`TicketDiscount`)/100), @frac_digits) AS `Amount`
          FROM `LTickets` WHERE `LTickets`.`IdHTicket`=OLD.`Id` AND `LTickets`.`lineType`=1 GROUP BY `LTickets`.`VATCode`) AS `V_VAT`
          SET `TotalVAT`.`Amount`=`TotalVAT`.`Amount`-`V_VAT`.`Amount`,
          `TotalVAT`.`Base`=`TotalVAT`.`Base`-(ROUND((`V_VAT`.`Amount`*100)/(100+`VATPercent`),@frac_digits)),
          `TotalVAT`.`Tax`=`TotalVAT`.`Tax`-(`V_VAT`.`Amount` - ROUND((`V_VAT`.`Amount`*100)/(100+`VATPercent`),@frac_digits))
          WHERE `Date`=DATE(OLD.`DatTim`) AND `NetStat`=OLD.`NetStat` AND `Code`=`V_VAT`.`VATCode`;

        UPDATE `TotalPayTickets`,
          (SELECT `PayMethod`, SUM(`Delivered`) AS `Delivered` FROM `PayTickets` WHERE `PayTickets`.`IdHTicket`=OLD.`Id` GROUP BY `PayMethod`) AS `V_PT`
          SET `TotalPayTickets`.`Delivered`=`TotalPayTickets`.`Delivered`-`V_PT`.`Delivered`
          WHERE `Date`=DATE(OLD.`DatTim`) AND `NetStat`=OLD.`NetStat` AND `TotalPayTickets`.`PayMethod`=`V_PT`.`PayMethod`;

        UPDATE `TotalScaleFamilies`,
          (SELECT CONCAT(YEAR(NEW.`DatTim`),'-',MONTH(NEW.`DatTim`), '-01') AS `fdt`,
          `Family`, SUM(IF(`Saleform`=1,`Weight`,0)) AS `fwght`, SUM(IF(`Saleform`=0,`Weight`,0)) AS `funits` FROM `LTickets` WHERE (`IdHTicket`=NEW.`Id` AND `LTickets`.`lineType`=1) GROUP BY `Family`) AS `VT_F`
          SET `TotalWeight`=`TotalWeight`-`VT_F`.`fwght` , `TotalUnits`=`TotalUnits`+ `VT_F`.`funits` WHERE (`Date`=`VT_F`.`fdt` AND `Scale`=NEW.`Scale` AND `TotalScaleFamilies`.`Family`=`VT_F`.`Family`);
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
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `invoicedetails`
--

LOCK TABLES `invoicedetails` WRITE;
/*!40000 ALTER TABLE `invoicedetails` DISABLE KEYS */;
INSERT INTO `invoicedetails` VALUES (1,1,1),(2,2,2),(3,3,3);
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
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `invoices`
--

LOCK TABLES `invoices` WRITE;
/*!40000 ALTER TABLE `invoices` DISABLE KEYS */;
INSERT INTO `invoices` VALUES (1,'160010100107','','2016-10-28','','0'),(2,'160010100108','','2016-10-28','','0'),(3,'160010100109','','2016-10-28','','0');
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
	UPDATE `HTickets` SET `Invoice`='', InvoiceDate=DATE(0) WHERE Id IN (SELECT IdHTicket FROM `InvoiceDetails` WHERE IdInvoice=OLD.`Id`);
	DELETE FROM `InvoiceDetails` WHERE IdInvoice=OLD.`Id`;
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
  `Tare` decimal(5,3) DEFAULT '0.000',
  `Label` varchar(64) DEFAULT '',
  `Nutrition` int(11) DEFAULT '0',
  `EAN13Print` varchar(12) DEFAULT '',
  `PacketExt` char(1) NOT NULL DEFAULT '-',
  `IsFrozen` tinyint(1) DEFAULT '0',
  `Icon` varchar(127) DEFAULT '',
  `Icon2` varchar(127) DEFAULT '',
  `QLevel` varchar(16) DEFAULT '',
  `Mixed` smallint(6) DEFAULT '0',
  `MixedWeight` decimal(5,3) DEFAULT '0.000',
  `CommonCard` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Code` (`Code`),
  KEY `Name` (`Name`(19)),
  KEY `ide_fmor` (`Family`,`FamilyOrder`)
) ENGINE=InnoDB AUTO_INCREMENT=442 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `items`
--

LOCK TABLES `items` WRITE;
/*!40000 ALTER TABLE `items` DISABLE KEYS */;
INSERT INTO `items` VALUES (1,0,'','',NULL,NULL,NULL,NULL,NULL,NULL,0,0,999,0,0,0,0,'',0.00,0,0.00,1,0,0,0,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(2,13720,'Limas','Limas','','','','','','',1,0,1,1,0,0,0,'',3.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13720.png','','',0,0.000,0),(3,13721,'Limones','Limones','','','','','','',1,0,2,1,0,0,0,'',2.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13721.png','','',0,0.000,0),(4,13722,'mandarina pequeña','mandarina pequeña ','','','','','','',1,0,6,1,0,0,0,'',1.80,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_100.png','','',0,0.000,0),(5,13723,'Naranjas Zumo','Naranjas Zumo','','','','','','',1,0,11,1,0,0,0,'',3.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752340345.jpg','','',0,0.000,0),(6,13724,'Naranjas','Naranjas','','','','','','',1,0,10,1,0,0,0,'',3.25,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_023.png','','',0,0.000,0),(7,13725,'Pomelos','Pomelos','','','','','','',1,0,12,1,0,0,0,'',2.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13725.png','','',0,0.000,0),(8,13726,'albericoc 1€','albericoc 1€','','','','','','',3,0,2,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400676785505.jpg','','',0,0.000,0),(9,13727,'Granada','Granada','','','','','','',2,0,9,1,0,0,0,'',2.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13727.png','','',0,0.000,0),(10,13728,'golden','golden','','','','','','',2,0,7,1,0,0,0,'',1.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13728.png','','',0,0.000,0),(11,13729,'Manzana Royal Gala','Manzana Royal Gala','','','','','','',2,0,14,1,0,0,0,'',2.30,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13729_1.png','','',0,0.000,0),(12,13731,'gabacho','gabacho','','','','','','',2,0,6,1,0,0,0,'',2.25,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_077.png','','',0,0.000,0),(13,13732,'melocoton de agua','melocoton de agua','','','','','','',2,0,16,1,0,0,0,'',1.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_079.png','','',0,0.000,0),(14,13733,'melocoton Calanda','melocoton Calanda','','','','','','',2,0,15,1,0,0,0,'',3.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_076.png','','',0,0.000,0),(15,13734,'Melón bruño','Melón bruño','','','','','','',2,0,19,1,0,0,0,'',2.20,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_092.png','','',0,0.000,0),(16,13735,'Melón Galia','Melón Galia','','','','','','',2,0,21,1,0,0,0,'',1.80,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_081.png','','',0,0.000,0),(17,13736,'melon bollo','melon bollo','','','','','','',2,0,18,1,0,0,0,'',2.40,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_092.png','','',0,0.000,0),(18,13737,'Membrillo','Membrillo','','','','','','',2,0,23,1,0,0,0,'',1.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_093.png','','',0,0.000,0),(19,13738,'Nectarina Amarilla','Nectarina Amarilla','','','','','','',2,0,24,1,0,0,0,'',2.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_104.png','','',0,0.000,0),(20,13739,'Nectarina Blanca','Nectarina Blanca','','','','','','',2,0,25,1,0,0,0,'',0.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'NECTARINA ROJA.jpg','','',0,0.000,0),(21,13740,'Pera Blanquilla','Pera Blanquilla','','','','','','',2,0,27,1,0,0,0,'',2.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_122.png','','',0,0.000,0),(22,13741,'Pera Conferencia','Pera Conferencia','','','','','','',2,0,28,1,0,0,0,'',2.65,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_126.png','','',0,0.000,0),(23,13742,'Pera ercolina','Pera ercolina','','','','','','',2,0,29,1,0,0,0,'',3.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_140.png','','',0,0.000,0),(24,13743,'Piña del monte','Piña del monte','','','','','','',4,0,19,1,0,0,0,'',2.25,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'10009_1.jpg','','',0,0.000,0),(25,13744,'Plátano canario','Plátano canario','','','','','','',2,0,34,1,0,0,0,'',2.25,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13744_1.png','','',0,0.000,0),(26,13745,'sandia rallada','sandia rallada','','','','','','',2,0,37,1,0,0,0,'',0.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_163.png','','',0,0.000,0),(27,13746,'Sandía negra','Sandía negra','','','','','','',2,0,36,1,0,0,0,'',1.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_159.png','','',0,0.000,0),(28,13747,'Castañas','Castañas','','','','','','',5,0,6,1,0,0,0,'',6.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_012_1.png','','',0,0.000,0),(29,13748,'Cerezas granel','Cerezas granel','','','','','','',3,0,11,1,0,0,0,'',6.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13748.png','','',0,0.000,0),(30,13749,'Ciruelas Amarillas','Ciruelas Amarillas','','','','','','',3,0,14,1,0,0,0,'',3.80,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13749.png','','',0,0.000,0),(31,13750,'Ciruelas Rojas','Ciruelas Rojas','','','','','','',3,0,15,1,0,0,0,'',3.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13750.png','','',0,0.000,0),(32,13751,'Frambuesa','Frambuesa','','','','','','',3,0,20,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13758.png','','',0,0.000,0),(33,13752,'Fresita','Fresita','','','','','','',3,0,21,1,0,0,0,'',4.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400075415061.jpg','','',0,0.000,0),(34,13753,'freson bandeja','freson bandeja','','','','','','',3,0,23,1,0,0,0,'',2.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'10011_1.jpg','','',0,0.000,0),(35,13754,'Grosella','Grosella','','','','','','',3,0,29,1,0,0,0,'',3.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13754.png','','',0,0.000,0),(36,13756,'Higos','Higos','','','','','','',3,0,31,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_036.png','','',0,0.000,0),(37,13757,'Kiwis Asturias','Kiwis Asturias','','','','','','',3,0,33,1,0,0,0,'',3.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_043.png','','',0,0.000,0),(38,13758,'Moras','Moras','','','','','','',3,0,38,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1386770865228.jpg','','',0,0.000,0),(39,13759,'uva blanca','uva blanca','','','','','','',3,0,45,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_167.png','','',0,0.000,0),(40,13760,'Aguacate','Aguacate','','','','','','',4,0,1,1,0,0,0,'',2.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13760.png','','',0,0.000,0),(41,13762,'Bananito','Bananito','','','','','','',4,0,5,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1382203892697.jpg','','',0,0.000,0),(42,13763,'Chirimoya','Chirimoya','','','','','','',4,0,7,1,0,0,0,'',2.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_016.png','','',0,0.000,0),(43,13764,'Coco','Coco','','','','','','',4,0,8,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_024.png','','',0,0.000,0),(44,13765,'Lichis','Lichis','','','','','','',4,0,11,1,0,0,0,'',2.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13765.png','','',0,0.000,0),(45,13766,'Mango','Mango','','','','','','',4,0,12,1,0,0,0,'',2.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_054.png','','',0,0.000,0),(46,13769,'Papaya','Papaya','','','','','','',4,0,15,1,0,0,0,'',4.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_110.png','','',0,0.000,0),(47,13773,'Almendras','Almendras','','','','','','',5,0,2,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'FS_004.png','','',0,0.000,0),(48,13774,'Avellanas','Avellanas','','','','','','',5,0,4,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(49,13775,'Cacahuetes','Cacahuetes','','','','','','',5,0,5,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'10007.jpg','','',0,0.000,0),(50,13776,'Dátiles','Dátiles','','','','','','',5,0,7,1,0,0,0,'',4.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'F_026.png','','',0,0.000,0),(51,13777,'Higos secos','Higos secos','','','','','','',5,0,10,1,0,0,0,'',2.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'F_041.png','','',0,0.000,0),(52,13778,'Maíz tostado','Maíz tostado','','','','','','',5,0,12,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'13778.png','','',0,0.000,0),(53,13779,'Mezcla frutos secos','Mezcla frutos secos','','','','','','',5,0,13,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(54,13780,'Nueces bolsa','Nueces bolsa','','','','','','',5,0,16,1,0,0,0,'',3.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'FS_019.png','','',0,0.000,0),(55,13781,'Orejones','Orejones','','','','','','',5,0,20,1,0,0,0,'',2.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'13781.png','','',0,0.000,0),(56,13782,'Pasas','Pasas','','','','','','',5,0,21,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'F_115.png','','',0,0.000,0),(57,13783,'Pipas','Pipas','','','','','','',5,0,25,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'13783.png','','',0,0.000,0),(58,13784,'Pistachos','Pistachos','','','','','','',5,0,26,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'FS_021.png','','',0,0.000,0),(59,13785,'Piñones','Piñones','','','','','','',5,0,24,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'FS_022.png','','',0,0.000,0),(60,13786,'Semillas calabaza','Semillas calabaza','','','','','','',5,0,27,1,0,0,0,'',0.00,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(61,13787,'Ajos Morados','Ajos Morados','','','','','','',6,0,2,1,0,0,0,'',5.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13788_1.png','','',0,0.000,0),(62,13789,'Apio','Apio','','','','','','',9,0,5,1,0,0,0,'',1.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13789.png','','',0,0.000,0),(63,13790,'Berenjenas','Berenjenas','','','','','','',6,0,6,1,0,0,0,'',1.40,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13790.png','','',0,0.000,0),(64,13791,'Calabacin','Calabacin','','','','','','',6,0,10,1,0,0,0,'',1.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13791.png','','',0,0.000,0),(65,13792,'Calabazas','Calabazas','','','','','','',6,0,12,1,0,0,0,'',1.30,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'10013.jpg','','',0,0.000,0),(66,13793,'Cardo','Cardo','','','','','','',9,0,13,1,0,0,0,'',0.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13793.png','','',0,0.000,0),(67,13794,'Cebollas Fuentes','Cebollas Fuentes','','','','','','',6,0,15,1,0,0,0,'',2.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13794.png','','',0,0.000,0),(68,13795,'Cebollas Figueras','Cebollas Figueras','','','','','','',6,0,14,1,0,0,0,'',1.10,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13796_1.png','','',0,0.000,0),(69,13797,'Champiñones granel','Champiñones granel','','','','','','',465,0,7,1,0,0,0,'',2.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400075430586.jpg','','',0,0.000,0),(70,13798,'esparrago roca','esparrago roca','','','','','','',9,0,23,1,0,0,0,'',3.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_044_1.png','','',0,0.000,0),(71,13799,'Nabos','Nabos','','','','','','',6,0,27,1,0,0,0,'',1.80,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13799.png','','',0,0.000,0),(72,13800,'Patatas Rojas','Patatas Rojas','','','','','','',6,0,35,1,0,0,0,'',1.20,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13800.png','','',0,0.000,0),(73,13801,'Patatas blanca','Patatas blanca','','','','','','',6,0,33,1,0,0,0,'',1.10,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'Patatatemprana.jpg','','',0,0.000,0),(74,13802,'Pepinos','Pepinos','','','','','','',6,0,38,1,0,0,0,'',1.30,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_076_1.png','','',0,0.000,0),(75,13803,'Pimiento verde','Pimiento verde','','','','','','',6,0,43,1,0,0,0,'',2.20,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_082.png','','',0,0.000,0),(76,13804,'Pimiento Rojo','Pimiento Rojo','','','','','','',6,0,41,1,0,0,0,'',2.65,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_085.png','','',0,0.000,0),(77,13806,'hierbas caldo','hierbas caldo','','','','','','',466,0,13,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13806_1.png','','',0,0.000,0),(78,13807,'Puerros','Puerros','','','','','','',9,0,36,1,0,0,0,'',2.20,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13807.png','','',0,0.000,0),(79,13808,'Rabanos redondo','Rabanos redondo','','','','','','',6,0,45,1,0,0,0,'',0.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_092.png','','',0,0.000,0),(80,13809,'Remolacha','Remolacha','','','','','','',6,0,46,1,0,0,0,'',1.60,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_093.png','','',0,0.000,0),(81,13810,'rovellons trozos','rovellons trozos','','','','','','',465,0,20,1,0,0,0,'',3.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'Rovellons-níscalos_2.jpg','','',0,0.000,0),(82,13811,'cherry Santa Susana','cherry Santa Susana','','','','','','',6,0,20,1,0,0,0,'',6.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1392819908196.jpg','','',0,0.000,0),(83,13813,'Tomates Rama','Tomates Rama','','','','','','',6,0,58,1,0,0,0,'',2.20,0,0.00,1,0,1,4,'',0,0,0.000,'',0,'','-',0,'13813.png','','',0,0.000,0),(84,13814,'Tomates Verdes','Tomates Verdes','','','','','','',6,0,59,1,0,0,0,'',2.40,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13814.png','','',0,0.000,0),(85,13815,'Zanahorias manojo','Zanahorias manojo','','','','','','',6,0,64,1,0,0,0,'',1.75,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'ZANAHORIAS RAMA.jpg','','',0,0.000,0),(86,13816,'Acelgas','Acelgas','','','','','','',9,0,1,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13816.png','','',0,0.000,0),(87,13817,'Alcachofas','Alcachofas','','','','','','',9,0,4,1,0,0,0,'',3.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13817_1.png','','',0,0.000,0),(88,13818,'Berros','Berros','','','','','','',463,0,4,1,0,0,0,'',1.99,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'BERROS.JPG','','',0,0.000,0),(89,13820,'Brocoli','Brocoli','','','','','','',9,0,9,1,0,0,0,'',1.10,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'10003.jpg','','',0,0.000,0),(90,13821,'Canónigos','Canónigos','','','','','','',464,0,5,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'RUCULA SELVATICA.JPG','','',0,0.000,0),(91,13822,'Cebollas Tiernas','Cebollas Tiernas','','','','','','',9,0,14,1,0,0,0,'',1.30,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_032.png','','',0,0.000,0),(92,13823,'Cogollos de 3','Cogollos de 3','','','','','','',464,0,7,1,0,0,0,'',1.65,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'COGOLLOS DE 2_2.JPG','','',0,0.000,0),(93,13824,'Coles De Bruselas','Coles De Bruselas','','','','','','',9,0,18,1,0,0,0,'',1.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_037.png','','',0,0.000,0),(94,13825,'Coliflor','Coliflor','','','','','','',9,0,19,1,0,0,0,'',2.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13825.png','','',0,0.000,0),(95,13826,'Endivias','Endivias','','','','','','',464,0,9,1,0,0,0,'',1.75,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'ENDIVIAS.jpg','','',0,0.000,0),(96,13827,'Escarola','Escarola','','','','','','',464,0,11,1,0,0,0,'',2.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'ESCAROLAS.jpg','','',0,0.000,0),(97,13828,'Espinacas','Espinacas','','','','','','',9,0,27,1,0,0,0,'',1.60,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13828.png','','',0,0.000,0),(98,13829,'Guisantes','Guisantes','','','','','','',9,0,28,1,0,0,0,'',3.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13833.png','','',0,0.000,0),(99,13830,'Habas','Habas','','','','','','',9,0,29,1,0,0,0,'',2.40,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13830.png','','',0,0.000,0),(100,13831,'hinojo','hinojo','','','','','','',9,0,30,1,0,0,0,'',2.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13831.png','','',0,0.000,0),(101,13832,'Judías Finas Kenia','Judías Finas Kenia','','','','','','',9,0,33,1,0,0,0,'',9.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'judiabobby.jpg','','',0,0.000,0),(102,13833,'Judías perona','Judías perona','','','','','','',9,0,34,1,0,0,0,'',6.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'Judiafina.jpg','','',0,0.000,0),(103,13834,'Lechuga Iceberg','Lechuga Iceberg','','','','','','',464,0,16,1,0,0,0,'',0.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13834.png','','',0,0.000,0),(104,13835,'Lechuga larga','Lechuga larga','','','','','','',464,0,17,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_062_1.png','','',0,0.000,0),(105,13836,'Lechuga Trocadero','Lechuga Trocadero','','','','','','',464,0,18,1,0,0,0,'',1.20,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_068.png','','',0,0.000,0),(106,13837,'Lombarda','Lombarda','','','','','','',9,0,35,1,0,0,0,'',1.80,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'Collombarda.jpg','','',0,0.000,0),(107,13839,'Perejil','Perejil','','','','','','',463,0,13,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13839.png','','',0,0.000,0),(108,13841,'Rúcula','Rúcula','','','','','','',464,0,23,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'RUCULA SELVATICA_1.JPG','','',0,0.000,0),(109,13842,'cocacola lata','cocacola lata','','','','','','',462,0,8,1,0,0,0,'',0.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'11901.png','','',0,0.000,0),(110,13844,'fanta naranja','fanta naranja','','','','','','',462,0,11,1,0,0,0,'',0.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'11905.png','','',0,0.000,0),(111,13846,'albericoc mirlo','albericoc mirlo','','','','','','',3,0,3,1,0,0,0,'',4.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400075406997.jpg','','',0,0.000,0),(112,13847,'ajo tierno','ajo tierno','','','','','','',9,0,2,1,0,0,0,'',1.80,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'Ajostiernos.jpg','','',0,0.000,0),(113,13848,'albahaca','albahaca','','','','','','',463,0,2,1,0,0,0,'',1.80,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'ALBAHACA ALFABREGA.JPG','','',0,0.000,0),(114,13849,'arandano','arandano','','','','','','',3,0,5,1,0,0,0,'',3.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'ARANDANOS.JPG','','',0,0.000,0),(115,13850,'boniato asado','boniato asado','','','','','','',6,0,8,1,0,0,0,'',3.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'batatas_1.jpg','','',0,0.000,0),(116,13851,'brevas','brevas','','','','','','',3,0,7,1,0,0,0,'',6.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400676617402.jpg','','',0,0.000,0),(117,13852,'broquil','broquil','','','','','','',9,0,10,1,0,0,0,'',2.40,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'ROMANESCU.jpg','','',0,0.000,0),(118,13853,'calçots','calçots','','','','','','',9,0,11,1,0,0,0,'',5.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'Calçots.jpg','','',0,0.000,0),(119,13854,'cebollita platillo','cebollita platillo','','','','','','',6,0,16,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752241310.jpg','','',0,0.000,0),(120,13855,'cebollino','cebollino','','','','','','',463,0,5,1,0,0,0,'',1.80,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'CEBOLLINO CIBULET.JPG','','',0,0.000,0),(121,13857,'ciruela mirabolat','ciruela mirabolat','','','','','','',3,0,13,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400577445820_1.jpg','','',0,0.000,0),(122,13858,'ciruela tolosa','ciruela tolosa','','','','','','',3,0,16,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'CIRUELA CLAUDIA.JPG','','',0,0.000,0),(123,13859,'claudia reina','claudia reina','','','','','','',3,0,17,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'CLAUDIA REINA_1.JPG','','',0,0.000,0),(124,13860,'cogollos de 6','cogollos de 6','','','','','','',464,0,8,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'COGOLLOS DE 6.JPG','','',0,0.000,0),(125,13861,'col','col','','','','','','',9,0,16,1,0,0,0,'',1.90,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_094.png','','',0,0.000,0),(126,13862,'chiribia','chiribia','','','','','','',6,0,21,1,0,0,0,'',2.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'Chirivia.jpg','','',0,0.000,0),(127,13863,'Esparrago triguero','Esparrago triguero','','','','','','',9,0,24,1,0,0,0,'',2.60,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_044_2.png','','',0,0.000,0),(128,13865,'girgola','girgola','','','','','','',465,0,11,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'Gírgolas.jpg','','',0,0.000,0),(129,13866,'guisantes bandeja','guisantes bandeja','','','','','','',466,0,11,1,0,0,0,'',3.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1392826244595.jpg','','',0,0.000,0),(130,13867,'granada bandeja','granada bandeja','','','','','','',466,0,10,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1380269643616.jpg','','',0,0.000,0),(131,13868,'habas peladas','habas peladas','','','','','','',466,0,12,1,0,0,0,'',2.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1392826235540.jpg','','',0,0.000,0),(132,13869,'higo blanco','higo blanco','','','','','','',3,0,30,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_042.png','','',0,0.000,0),(133,13870,'coll de dama','coll de dama','','','','','','',3,0,19,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_040.png','','',0,0.000,0),(134,13871,'hoja de roble','hoja de roble','','','','','','',464,0,15,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'HOJA DE ROBLE.JPG','','',0,0.000,0),(135,13872,'gengibre','gengibre','','','','','','',463,0,10,1,0,0,0,'',2.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_056.png','','',0,0.000,0),(136,13873,'kaqui bandeja','kaqui bandeja','','','','','','',2,0,10,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_011_1.png','','',0,0.000,0),(137,13875,'Mandarina','Mandarina','','','','','','',1,0,3,1,0,0,0,'',0.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13722.png','','',0,0.000,0),(138,13876,'manzana acida','manzana acida','','','','','','',2,0,12,1,0,0,0,'',2.20,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_061.png','','',0,0.000,0),(139,13877,'manzana fuji','manzana fuji','','','','','','',2,0,13,1,0,0,0,'',2.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'10001.jpg','','',0,0.000,0),(140,13878,'golden oferta','golden oferta','','','','','','',2,0,8,1,0,0,0,'',0.79,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13728.png','','',0,0.000,0),(141,13879,'pink lady','pink lady','','','','','','',2,0,32,1,0,0,0,'',3.25,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1384544081073.jpg','','',0,0.000,0),(142,13880,'paraguayo','paraguayo','','','','','','',2,0,26,1,0,0,0,'',2.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_113.png','','',0,0.000,0),(143,13881,'melon charentais','melon charentais','','','','','','',2,0,20,1,0,0,0,'',1.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_089.png','','',0,0.000,0),(144,13882,'menta','menta','','','','','','',463,0,12,1,0,0,0,'',1.00,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'MENTA.JPG','','',0,0.000,0),(145,13883,'sanguina bolsa','sanguina bolsa','','','','','','',1,0,14,1,0,0,0,'',2.90,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1390571966883.jpg','','',0,0.000,0),(146,13884,'nispero','nispero','','','','','','',3,0,40,1,0,0,0,'',4.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_108.png','','',0,0.000,0),(147,13885,'trumfos pirineu','trumfos pirineu','','','','','','',6,0,61,1,0,0,0,'',3.40,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1382368994716.jpg','','',0,0.000,0),(148,13886,'patata gallega','patata gallega','','','','','','',6,0,31,1,0,0,0,'',3.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1382368989766.jpg','','',0,0.000,0),(149,13887,'comice','comice','','','','','','',2,0,4,1,0,0,0,'',2.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1382372822498.jpg','','',0,0.000,0),(150,13889,'perita castells','perita castells','','','','','','',3,0,8,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_124.png','','',0,0.000,0),(151,13890,'pimiento padron','pimiento padron','','','','','','',6,0,40,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_083.png','','',0,0.000,0),(152,13891,'piña Costa Rica','piña Costa Rica','','','','','','',4,0,18,1,0,0,0,'',1.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'10009_1.jpg','','',0,0.000,0),(153,13893,'rabano largo','rabano largo','','','','','','',6,0,44,1,0,0,0,'',0.90,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1382448095041.jpg','','',0,0.000,0),(154,13894,'achicoria','achicoria','','','','','','',464,0,1,1,0,0,0,'',1.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(155,13895,'camagroc','camagroc','','','','','','',465,0,2,1,0,0,0,'',3.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752222821.jpg','','',0,0.000,0),(156,13896,'rovellons mediano','rovellons mediano','','','','','','',465,0,19,1,0,0,0,'',6.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'Rovellons-níscalos_1.jpg','','',0,0.000,0),(157,13897,'rovellons boton','rovellons boton','','','','','','',465,0,18,1,0,0,0,'',8.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'Rovellons-níscalos.jpg','','',0,0.000,0),(158,13898,'trompeta de la muerte','trompeta de la muerte','','','','','','',465,0,23,1,0,0,0,'',3.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1380789671176.jpg','','',0,0.000,0),(159,13899,'rosinyol','rosinyol','','','','','','',465,0,16,1,0,0,0,'',3.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'ROSIÑOL.JPG','','',0,0.000,0),(160,13900,'sharoni','sharoni','','','','','','',2,0,38,1,0,0,0,'',3.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'KAKI.jpg','','',0,0.000,0),(161,13901,'tomate kumato','tomate kumato','','','','','','',6,0,52,1,0,0,0,'',2.25,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1389710988091.jpg','','',0,0.000,0),(162,13902,'tomate colgar ristra','tomate colgar ristra','','','','','','',6,0,50,1,0,0,0,'',5.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'TOMATE RISTRA.JPG','','',0,0.000,0),(163,13903,'bandeja sucar','bandeja sucar','','','','','','',6,0,4,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'TOMATE COLGAR GRANEL.JPG','','',0,0.000,0),(164,13904,'tomate pera','tomate pera','','','','','','',6,0,55,1,0,0,0,'',2.30,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_112.png','','',0,0.000,0),(165,13905,'tomate montserrat','tomate montserrat','','','','','','',6,0,54,1,0,0,0,'',4.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752254825.jpg','','',0,0.000,0),(166,13906,'tomate cor de bou','tomate cor de bou','','','','','','',6,0,51,1,0,0,0,'',3.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'COR DE BOU.JPG','','',0,0.000,0),(167,13907,'tomate raf','tomate raf','','','','','','',6,0,57,1,0,0,0,'',4.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_107.png','','',0,0.000,0),(168,13908,'tomate raf bandeja','tomate raf bandeja','','','','','','',466,0,29,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1381172084217_1.jpg','','',0,0.000,0),(169,13909,'uva roja','uva roja','','','','','','',3,0,44,1,0,0,0,'',1.20,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_173.png','','',0,0.000,0),(170,13910,'uva moscatel','uva moscatel','','','','','','',3,0,43,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_169.png','','',0,0.000,0),(171,13911,'zanahoria','zanahoria','','','','','','',6,0,63,1,0,0,0,'',1.10,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_115.png','','',0,0.000,0),(172,13912,'laurel','laurel','','','','','','',463,0,11,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1380112062441.jpg','','',0,0.000,0),(173,13913,'romero','romero','','','','','','',463,0,14,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'ROMERO ROMANI.JPG','','',0,0.000,0),(174,13914,'tomillo farigola','tomillo farigola','','','','','','',463,0,15,1,0,0,0,'',1.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'TOMILLO FARIGOLA.JPG','','',0,0.000,0),(175,13915,'eneldo','eneldo','','','','','','',463,0,7,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'ENELDO.JPG','','',0,0.000,0),(176,13916,'cilantro','cilantro','','','','','','',463,0,6,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'CILANTRO.JPG','','',0,0.000,0),(177,13917,'tomate vasco','tomate vasco','','','','','','',6,0,60,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13814_1.png','','',0,0.000,0),(178,13918,'shii take','shii take','','','','','','',465,0,21,1,0,0,0,'',2.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'SHII TAKE.JPG','','',0,0.000,0),(179,13919,'mezclum bandeja','mezclum bandeja','','','','','','',464,0,21,1,0,0,0,'',1.99,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752282686.jpg','','',0,0.000,0),(180,13920,'naranja oferta','naranja oferta','','','','','','',1,0,9,1,0,0,0,'',1.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13723.png','','',0,0.000,0),(181,13921,'calabaza troceada','calabaza troceada','','','','','','',466,0,6,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'CALABAZA TROCEADA.JPG','','',0,0.000,0),(182,13922,'piña bandeja','piña bandeja','','','','','','',466,0,23,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'PIÑA BANDEJA.JPG','','',0,0.000,0),(183,13923,'macedonia','macedonia','','','','','','',466,0,15,1,0,0,0,'',2.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'MACEDONIA_1.JPG','','',0,0.000,0),(184,13925,'sandia bandeja','sandia bandeja','','','','','','',466,0,26,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'SANDIA BANDEJA_1.JPG','','',0,0.000,0),(185,13926,'menestra  verduras','menestra  verduras','','','','','','',466,0,18,1,0,0,0,'',3.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'MENESTRA VERDURAS.JPG','','',0,0.000,0),(186,13927,'acelga bandeja','acelga bandeja','','','','','','',466,0,1,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752213123.jpg','','',0,0.000,0),(187,13928,'agua medio','agua medio','','','','','','',462,0,6,1,0,0,0,'',0.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'aguaviladrau.png','','',0,0.000,0),(188,13929,'agua grande','agua grande','','','','','','',462,0,4,1,0,0,0,'',0.75,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'11906.png','','',0,0.000,0),(189,13931,'cerveza','cerveza','','','','','','',462,0,7,1,0,0,0,'',0.80,0,0.00,0,0,1,0,'',0,0,0.000,'',0,'','-',0,'ESTRELLA DAMM.JPG','','',0,0.000,0),(190,13932,'zero','zero','','','','','','',462,0,23,1,0,0,0,'',0.80,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'COCA COLA ZERO.JPG','','',0,0.000,0),(191,13933,'nestea','nestea','','','','','','',462,0,15,1,0,0,0,'',0.90,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'11916.png','','',0,0.000,0),(192,13938,'500','500','','','','','','',462,0,1,1,0,0,0,'',1.00,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(193,13939,'acuarius','acuarius','','','','','','',462,0,2,1,0,0,0,'',0.90,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(194,13940,'red bull','red bull','','','','','','',462,0,17,1,0,0,0,'',1.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'11922.png','','',0,0.000,0),(195,13941,'tonica','tonica','','','','','','',462,0,21,1,0,0,0,'',0.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'11923.png','','',0,0.000,0),(196,13942,'sprite','sprite','','','','','','',462,0,20,1,0,0,0,'',0.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'11924.png','','',0,0.000,0),(197,13943,'agua limon','agua limon','','','','','','',462,0,5,1,0,0,0,'',1.40,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752335731.jpg','','',0,0.000,0),(198,13944,'fanta limon','fanta limon','','','','','','',462,0,10,1,0,0,0,'',0.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'11904.png','','',0,0.000,0),(199,13945,'gourmet','gourmet','','','','','','',464,0,14,1,0,0,0,'',2.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'GOURMET DE LA HUERTA_1.JPG','','',0,0.000,0),(200,13947,'cogollos de 2','cogollos de 2','','','','','','',464,0,6,1,0,0,0,'',1.20,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'COGOLLOS DE 2_1.JPG','','',0,0.000,0),(201,13948,'bandeja de escarola','bandeja de escarola','','','','','','',464,0,2,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752248715.jpg','','',0,0.000,0),(202,13949,'pack 6','pack 6','','','','','','',462,0,16,1,0,0,0,'',3.90,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752328889.jpg','','',0,0.000,0),(203,13950,'bollo etiqueta negra','bollo etiqueta negra','','','','','','',2,0,2,1,0,0,0,'',2.80,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_092_1.png','','',0,0.000,0),(204,13953,'melon bandeja','melon bandeja','','','','','','',466,0,17,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'MELON BANDEJA.JPG','','',0,0.000,0),(205,13955,'kiwi gold','kiwi gold','','','','','','',3,0,32,1,0,0,0,'',6.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_044.png','','',0,0.000,0),(206,13956,'chumbo','chumbo','','','','','','',3,0,12,1,0,0,0,'',1.99,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752262081.jpg','','',0,0.000,0),(207,13958,'patata agria','patata agria','','','','','','',6,0,28,1,0,0,0,'',0.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'PATATA ROJA.jpg','','',0,0.000,0),(208,13959,'cebolla normal','cebolla normal','','','','','','',6,0,19,1,0,0,0,'',0.70,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'CEBOLLA DULCE_1.jpg','','',0,0.000,0),(209,13960,'ensalada brotes','ensalada brotes','','','','','','',464,0,10,1,0,0,0,'',1.99,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'BROTES TIERNOS_2.JPG','','',0,0.000,0),(210,13961,'ribes 5l','ribes 5l','','','','','','',462,0,18,1,0,0,0,'',1.35,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752206097.jpg','','',0,0.000,0),(211,13962,'chumbo pelado','chumbo pelado','','','','','','',466,0,7,1,0,0,0,'',2.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752274913.jpg','','',0,0.000,0),(212,13963,'cherry oferta','cherry oferta','','','','','','',6,0,18,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_104.png','','',0,0.000,0),(213,13964,'gazpacho','gazpacho','','','','','','',462,0,12,1,0,0,0,'',3.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752320752.jpg','','',0,0.000,0),(214,13965,'garbanzo grande','garbanzo grande','','','','','','',467,0,29,1,0,0,0,'',3.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752230133.jpg','','',0,0.000,0),(215,13966,'alubia grande','alubia grande','','','','','','',467,0,8,1,0,0,0,'',3.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752311293.jpg','','',0,0.000,0),(216,13967,'lenteja grande','lenteja grande','','','','','','',467,0,33,1,0,0,0,'',3.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(217,13968,'garbanzo pequeño','garbanzo pequeño','','','','','','',467,0,30,1,0,0,0,'',2.65,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752230133_1.jpg','','',0,0.000,0),(218,13969,'alubia pequeña','alubia pequeña','','','','','','',467,0,9,1,0,0,0,'',2.65,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752311293_1.jpg','','',0,0.000,0),(219,13970,'lenteja pequeño','lenteja pequeño','','','','','','',467,0,34,1,0,0,0,'',2.65,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(220,13971,'crema alcachofa','crema alcachofa','','','','','','',467,0,18,1,0,0,0,'',3.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429802752.jpg','','',0,0.000,0),(221,13972,'crema zanahoria','crema zanahoria','','','','','','',467,0,25,1,0,0,0,'',2.99,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429798220.jpg','','',0,0.000,0),(222,13973,'crema calabacin','crema calabacin','','','','','','',467,0,20,1,0,0,0,'',2.99,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429744249.jpg','','',0,0.000,0),(223,13974,'crema verduras','crema verduras','','','','','','',467,0,23,1,0,0,0,'',2.99,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429795590.jpg','','',0,0.000,0),(224,13975,'crema esparragos','crema esparragos','','','','','','',467,0,22,1,0,0,0,'',3.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429788979.jpg','','',0,0.000,0),(225,13976,'crema calabaza','crema calabaza','','','','','','',467,0,21,1,0,0,0,'',2.99,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429786505.jpg','','',0,0.000,0),(226,13977,'crema vichyssoise','crema vichyssoise','','','','','','',467,0,24,1,0,0,0,'',2.99,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429751629.jpg','','',0,0.000,0),(227,13978,'crema boletus','crema boletus','','','','','','',467,0,19,1,0,0,0,'',2.99,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429768664.jpg','','',0,0.000,0),(228,13979,'caldo verduras','caldo verduras','','','','','','',467,0,14,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429764997.jpg','','',0,0.000,0),(229,13980,'caldo pollo','caldo pollo','','','','','','',467,0,13,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429758228.jpg','','',0,0.000,0),(230,13981,'caldo pescado','caldo pescado','','','','','','',467,0,12,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429778721.jpg','','',0,0.000,0),(231,13982,'tomate frito extra','tomate frito extra','','','','','','',467,0,55,1,0,0,0,'',1.80,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(232,13983,'habitas baby','habitas baby','','','','','','',467,0,31,1,0,0,0,'',4.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379429774276.jpg','','',0,0.000,0),(233,13984,'bonito','bonito','','','','','','',467,0,10,1,0,0,0,'',2.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379430305316.jpg','','',0,0.000,0),(234,13985,'6 huevos num. 1','6 huevos num. 1','','','','','','',467,0,2,1,0,0,0,'',2.05,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379430312200_2.jpg','','',0,0.000,0),(235,13986,'6 huevos num. 3','6 huevos num. 3','','','','','','',467,0,3,1,0,0,0,'',1.35,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379430309770.jpg','','',0,0.000,0),(236,13987,'bonito oferta','bonito oferta','','','','','','',467,0,11,1,0,0,0,'',6.00,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1379433331987.jpg','','',0,0.000,0),(237,13989,'miel mil flores 500g','miel mil flores 500g','','','','','','',467,0,43,1,0,0,0,'',5.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(238,13990,'miel mil flores 250g','miel mil flores 250g','','','','','','',467,0,42,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(239,13991,'miel romero','miel romero','','','','','','',467,0,44,1,0,0,0,'',6.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(240,13992,'miel tomillo','miel tomillo','','','','','','',467,0,45,1,0,0,0,'',6.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(241,13993,'mandarina villa maria','mandarina villa maria','','','','','','',1,0,4,1,0,0,0,'',3.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13722.png','','',0,0.000,0),(242,13994,'ñora','ñora','','','','','','',5,0,19,1,0,0,0,'',0.40,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_179.png','','',0,0.000,0),(243,13995,'guindilla','guindilla','','','','','','',5,0,9,1,0,0,0,'',0.40,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_078.png','','',0,0.000,0),(244,13996,'tomate pimiento','tomate pimiento','','','','','','',6,0,56,1,0,0,0,'',2.99,0,0.00,1,0,1,0,'',0,0,0.000,'',0,'','-',0,'1381340299118.jpg','','',0,0.000,0),(245,13997,'tomate barbastro','tomate barbastro','','','','','','',6,0,49,1,0,0,0,'',3.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1389710990576.jpg','','',0,0.000,0),(246,13998,'ajo y perejil','ajo y perejil','','','','','','',463,0,1,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1381333760065.jpg','','',0,0.000,0),(247,13999,'salsa alcachofa','salsa alcachofa','','','','','','',467,0,51,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(248,14000,'salsa piquillo','salsa piquillo','','','','','','',467,0,54,1,0,0,0,'',1.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(249,14001,'salsa esparragos','salsa esparragos','','','','','','',467,0,53,1,0,0,0,'',1.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(250,14002,'madroños','madroños','','','','','','',3,0,35,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400676706723_1.jpg','','',0,0.000,0),(251,14003,'escalonia','escalonia','','','','','','',6,0,22,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1382368046648.jpg','','',0,0.000,0),(252,14005,'mermelada','mermelada','','','','','','',467,0,36,1,0,0,0,'',3.65,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(253,14006,'mermelada especial','mermelada especial','','','','','','',467,0,37,1,0,0,0,'',4.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(254,14008,'boniato blanco','boniato blanco','','','','','','',6,0,9,1,0,0,0,'',1.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1383752469545.jpg','','',0,0.000,0),(255,14010,'verde doncella','verde doncella','','','','','','',2,0,39,1,0,0,0,'',3.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1383752472836.jpg','','',0,0.000,0),(256,14011,'pimiento choricero','pimiento choricero','','','','','','',5,0,23,1,0,0,0,'',0.40,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1383752480396.jpg','','',0,0.000,0),(257,14012,'persimon','persimon','','','','','','',2,0,31,1,0,0,0,'',1.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1383752476673.jpg','','',0,0.000,0),(258,14013,'escudella','escudella','','','','','','',467,0,27,1,0,0,0,'',3.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1379429758228_1.jpg','','',0,0.000,0),(259,14014,'patatas mini','patatas mini','','','','','','',6,0,34,1,0,0,0,'',3.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1389716199292.jpg','','',0,0.000,0),(260,14015,'mongeta ganxet','mongeta ganxet','','','','','','',5,0,15,1,0,0,0,'',7.90,0,0.00,1,0,1,1,'',0,0,0.000,'',0,'','-',0,'1384544084241.jpg','','',0,0.000,0),(261,14016,'fredolic','fredolic','','','','','','',465,0,10,1,0,0,0,'',2.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1384788459067.jpg','','',0,0.000,0),(262,14017,'aguacate coktail','aguacate coktail','','','','','','',4,0,3,1,0,0,0,'',4.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1384544161037_1.jpg','','',0,0.000,0),(263,14019,'llanega','llanega','','','','','','',465,0,12,1,0,0,0,'',19.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1385053947659.jpg','','',0,0.000,0),(264,14020,'remenat setas','remenat setas','','','','','','',465,0,15,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1386457581801.jpg','','',0,0.000,0),(265,14023,'physalis','physalis','','','','','','',4,0,16,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1386770861872.jpg','','',0,0.000,0),(266,14025,'trina','trina','','','','','','',462,0,22,1,0,0,0,'',0.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'trinanaranja.png','','',0,0.000,0),(267,14026,'uva sin pepitas','uva sin pepitas','','','','','','',3,0,46,1,0,0,0,'',4.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_171.png','','',0,0.000,0),(268,14027,'farcellet','farcellet','','','','','','',463,0,8,1,0,0,0,'',0.65,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1389717074258.jpg','','',0,0.000,0),(269,14028,'piña baby','piña baby','','','','','','',4,0,17,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_148.png','','',0,0.000,0),(270,14029,'comice pais','comice pais','','','','','','',2,0,5,1,0,0,0,'',2.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1388243192874.jpg','','',0,0.000,0),(271,14030,'bolsita 12 uvas','bolsita 12 uvas','','','','','','',466,0,4,1,0,0,0,'',0.75,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(272,14031,'mermelada tomate','mermelada tomate','','','','','','',467,0,38,1,0,0,0,'',5.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(273,14032,'maracuya','maracuya','','','','','','',4,0,14,1,0,0,0,'',9.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13768.png','','',0,0.000,0),(274,14033,'feo de Tudela','feo de Tudela','','','','','','',6,0,23,1,0,0,0,'',4.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1381340302835.jpg','','',0,0.000,0),(275,14034,'freson maresme','freson maresme','','','','','','',3,0,26,1,0,0,0,'',4.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1389793073900.jpg','','',0,0.000,0),(276,14035,'perona cuarto kg','perona cuarto kg','','','','','','',466,0,21,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1389276863295.jpg','','',0,0.000,0),(277,14036,'judia fina bandeja','judia fina bandeja','','','','','','',466,0,14,1,0,0,0,'',2.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1386457573287.jpg','','',0,0.000,0),(278,14037,'salsa calçots','salsa calçots','','','','','','',467,0,52,1,0,0,0,'',1.99,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(279,14038,'tirabeques','tirabeques','','','','','','',9,0,38,1,0,0,0,'',11.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_101_1.png','','',0,0.000,0),(280,14039,'champiñon portobello','champiñon portobello','','','','','','',465,0,9,1,0,0,0,'',1.60,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1390317185406.jpg','','',0,0.000,0),(281,14040,'espinaca bolsa','espinaca bolsa','','','','','','',9,0,26,1,0,0,0,'',1.80,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1390317188661.jpg','','',0,0.000,0),(282,14041,'perona medio kg','perona medio kg','','','','','','',466,0,22,1,0,0,0,'',3.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1390387890019.jpg','','',0,0.000,0),(283,14042,'sanguina','sanguina','','','','','','',1,0,13,1,0,0,0,'',2.20,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1390387887190.jpg','','',0,0.000,0),(284,14043,'fruta de la pasion','fruta de la pasion','','','','','','',4,0,9,1,0,0,0,'',9.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1389623621041.jpg','','',0,0.000,0),(285,14044,'mandarina oferta','mandarina oferta','','','','','','',1,0,5,1,0,0,0,'',1.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_100.png','','',0,0.000,0),(286,14045,'nuez pacana','nuez pacana','','','','','','',5,0,18,1,0,0,0,'',8.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'FS_020.png','','',0,0.000,0),(287,14046,'maravilla','maravilla','','','','','','',464,0,20,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_041.png','','',0,0.000,0),(288,14047,'lamullo verde','lamullo verde','','','','','','',6,0,24,1,0,0,0,'',1.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_089.png','','',0,0.000,0),(289,14048,'nata','nata','','','','','','',3,0,39,1,0,0,0,'',2.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1392380811788.jpg','','',0,0.000,0),(290,14050,'freson caja 2kg','freson caja 2kg','','','','','','',3,0,25,1,0,0,0,'',3.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'10011_3.jpg','','',0,0.000,0),(291,14051,'cherry bola rama','cherry bola rama','','','','','','',6,0,17,1,0,0,0,'',3.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1392818791793.jpg','','',0,0.000,0),(292,14052,'tirabeque pais','tirabeque pais','','','','','','',9,0,37,1,0,0,0,'',0.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_101.png','','',0,0.000,0),(293,14053,'patata violeta','patata violeta','','','','','','',6,0,36,1,0,0,0,'',4.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1396013856396.jpg','','',0,0.000,0),(294,14054,'papaya bandeja','papaya bandeja','','','','','','',466,0,20,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1396013794046.jpg','','',0,0.000,0),(295,14055,'calabacin redondo','calabacin redondo','','','','','','',6,0,11,1,0,0,0,'',1.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1396528846922.jpg','','',0,0.000,0),(296,14056,'lenteja pardina','lenteja pardina','','','','','','',5,0,11,1,0,0,0,'',1.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'FS_012.png','','',0,0.000,0),(297,14057,'garbanzo','garbanzo','','','','','','',5,0,8,1,0,0,0,'',1.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'13437.png','','',0,0.000,0),(298,14058,'alubia pinta','alubia pinta','','','','','','',5,0,3,1,0,0,0,'',1.99,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'FS_013.png','','',0,0.000,0),(299,14061,'kiwi zespri','kiwi zespri','','','','','','',3,0,34,1,0,0,0,'',3.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13757.png','','',0,0.000,0),(300,14062,'berengena rallada','berengena rallada','','','','','','',6,0,5,1,0,0,0,'',1.80,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400075410532.jpg','','',0,0.000,0),(301,14063,'salmorejo','salmorejo','','','','','','',462,0,19,1,0,0,0,'',2.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(302,14066,'zumo 1l berrys','zumo 1l berrys','','','','','','',462,0,24,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(303,14067,'albericoc orange','albericoc orange','','','','','','',3,0,4,1,0,0,0,'',2.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400075406997_1.jpg','','',0,0.000,0),(304,14069,'Cereza  bandeja','Cereza  bandeja','','','','','','',3,0,9,1,0,0,0,'',3.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13748.png','','',0,0.000,0),(305,14070,'champi bandeja','champi bandeja','','','','','','',465,0,6,1,0,0,0,'',1.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400075399571.jpg','','',0,0.000,0),(306,14071,'judia fina pais','judia fina pais','','','','','','',9,0,32,1,0,0,0,'',5.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_058.png','','',0,0.000,0),(307,14072,'esparrago corto','esparrago corto','','','','','','',467,0,28,1,0,0,0,'',1.80,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(308,14073,'almendra fresca','almendra fresca','','','','','','',5,0,1,1,0,0,0,'',2.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(309,14074,'moniqui bandeja','moniqui bandeja','','','','','','',3,0,37,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400676785505_1.jpg','','',0,0.000,0),(310,14075,'bollo bodega','bollo bodega','','','','','','',2,0,1,1,0,0,0,'',3.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_092.png','','',0,0.000,0),(311,14076,'pimiento tricolor','pimiento tricolor','','','','','','',6,0,42,1,0,0,0,'',1.80,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'10008.jpg','','',0,0.000,0),(312,14078,'platano bandeja','platano bandeja','','','','','','',2,0,33,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13744.png','','',0,0.000,0),(313,14080,'melon villa maria','melon villa maria','','','','','','',2,0,22,1,0,0,0,'',2.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_092.png','','',0,0.000,0),(314,14081,'pera forelle','pera forelle','','','','','','',2,0,30,1,0,0,0,'',2.80,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_134.png','','',0,0.000,0),(315,14082,'mezclum bolsa','mezclum bolsa','','','','','','',464,0,22,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1378752282686_1.jpg','','',0,0.000,0),(316,14083,'Rosa de Hijar','Rosa de Hijar','','','','','','',6,0,47,1,0,0,0,'',3.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1381340302835_2.jpg','','',0,0.000,0),(317,14084,'claudia reina oferta','claudia reina oferta','','','','','','',3,0,18,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'CLAUDIA REINA_2.JPG','','',0,0.000,0),(318,14085,'manzana reineta','manzana reineta','','','','','','',2,0,35,1,0,0,0,'',2.25,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_071.png','','',0,0.000,0),(319,14086,'borrajas','borrajas','','','','','','',9,0,8,1,0,0,0,'',1.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_017.png','','',0,0.000,0),(320,14087,'borraja bandeja','borraja bandeja','','','','','','',466,0,5,1,0,0,0,'',2.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13819.png','','',0,0.000,0),(321,14088,'mandarina verde','mandarina verde','','','','','','',1,0,7,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13720_1.png','','',0,0.000,0),(322,14090,'guayaba','guayaba','','','','','','',4,0,10,1,0,0,0,'',4.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13761.png','','',0,0.000,0),(323,14091,'torres artesanal','torres artesanal','','','','','','',467,0,56,1,0,0,0,'',2.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(324,14092,'torres trufa','torres trufa','','','','','','',467,0,57,1,0,0,0,'',2.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(325,14094,'boniato','boniato','','','','','','',6,0,7,1,0,0,0,'',1.20,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'batatas.jpg','','',0,0.000,0),(326,14096,'ceps (boletus)','ceps (boletus)','','','','','','',465,0,4,1,0,0,0,'',30.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'BOLETUS.jpg','','',0,0.000,0),(327,14098,'champiñon laminado','champiñon laminado','','','','','','',465,0,8,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13797.png','','',0,0.000,0),(328,14099,'miel de limon','miel de limon','','','','','','',467,0,41,1,0,0,0,'',6.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(329,14100,'miel de azahar','miel de azahar','','','','','','',467,0,39,1,0,0,0,'',6.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(330,14101,'miel de eucalipto','miel de eucalipto','','','','','','',467,0,40,1,0,0,0,'',7.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(331,14102,'empanadilla de la Carmen','empanadilla de la Carmen','','','','','','',466,0,8,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(332,14103,'brotes de soja','brotes de soja','','','','','','',464,0,3,1,0,0,0,'',1.80,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(333,14105,'cebolleta fina','cebolleta fina','','','','','','',9,0,15,1,0,0,0,'',1.20,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_031.png','','',0,0.000,0),(334,14106,'campestre','campestre','','','','','','',464,0,4,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'GOURMET DE LA HUERTA.JPG','','',0,0.000,0),(335,14107,'arandano rojo','arandano rojo','','','','','','',3,0,6,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'FS_010.png','','',0,0.000,0),(336,14108,'col de paperina','col de paperina','','','','','','',9,0,17,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_094_1.png','','',0,0.000,0),(337,14109,'moixerno','moixerno','','','','','','',465,0,13,1,0,0,0,'',2.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(338,14110,'azafran','azafran','','','','','','',463,0,3,1,0,0,0,'',2.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(339,14111,'picada','picada','','','','','','',5,0,22,1,0,0,0,'',0.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(340,14112,'mangostan','mangostan','','','','','','',4,0,13,1,0,0,0,'',15.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13767.png','','',0,0.000,0),(341,14113,'alcachofa brasa','alcachofa brasa','','','','','','',9,0,3,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13817.png','','',0,0.000,0),(342,14114,'flores','flores','','','','','','',463,0,9,1,0,0,0,'',6.90,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'V_071.png','','',0,0.000,0),(343,14115,'freson bandeja troceado','freson bandeja troceado','','','','','','',466,0,9,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(344,14117,'freson caja 1kg','freson caja 1kg','','','','','','',3,0,24,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'10011_2.jpg','','',0,0.000,0),(345,14118,'Piña Frunet','Piña Frunet','','','','','','',4,0,20,1,0,0,0,'',2.65,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'10009_1.jpg','','',0,0.000,0),(346,14119,'golosina','golosina','','','','','','',3,0,28,1,0,0,0,'',1.00,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(347,14120,'calçots a la leña','calçots a la leña','','','','','','',9,0,12,1,0,0,0,'',12.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'Calçots_1.jpg','','',0,0.000,0),(348,14121,'pimiento amarillo','pimiento amarillo','','','','','','',6,0,39,1,0,0,0,'',3.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_079.png','','',0,0.000,0),(349,14122,'piquillo en tiras','piquillo en tiras','','','','','','',467,0,49,1,0,0,0,'',2.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(350,14123,'escalivada pimiento+beren','escalivada pimiento+beren','','','','','','',467,0,26,1,0,0,0,'',4.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(351,14125,'choco crips','choco crips','','','','','','',467,0,16,1,0,0,0,'',1.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(352,14129,'Raf Pata Negra','Raf Pata Negra','','','','','','',466,0,25,1,0,0,0,'',2.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1380206928988.jpg','','',0,0.000,0),(353,14130,'esparrago bandeja','esparrago bandeja','','','','','','',9,0,21,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_044.png','','',0,0.000,0),(354,14132,'zumo eco 200ml','zumo eco 200ml','','','','','','',462,0,26,1,0,0,0,'',1.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(355,14133,'zumo eco pack 3','zumo eco pack 3','','','','','','',462,0,27,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(356,14134,'monterosa bandeja','monterosa bandeja','','','','','','',466,0,19,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1381340302835.jpg','','',0,0.000,0),(357,14135,'bandeja cor de bou','bandeja cor de bou','','','','','','',466,0,3,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'COR DE BOU.JPG','','',0,0.000,0),(358,14136,'nispero almibar','nispero almibar','','','','','','',467,0,48,1,0,0,0,'',4.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(359,14138,'fusion fresa-piña','fusion fresa-piña','','','','','','',3,0,27,1,0,0,0,'',4.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(360,14139,'ajo negro','ajo negro','','','','','','',6,0,1,1,0,0,0,'',5.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'13788.png','','',0,0.000,0),(361,14140,'PICOTA','PICOTA','','','','','','',3,0,42,1,0,0,0,'',3.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'10004_3.jpg','','',0,0.000,0),(362,14141,'aguacate bandeja','aguacate bandeja','','','','','','',4,0,2,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13760.png','','',0,0.000,0),(363,14142,'ketchup','ketchup','','','','','','',467,0,32,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(364,14143,'tomate monte rosa','tomate monte rosa','','','','','','',6,0,53,1,0,0,0,'',2.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'1381340302835_1.jpg','','',0,0.000,0),(365,14144,'barreja bolets deshidratada','barreja bolets deshidratada','','','','','','',465,0,1,1,0,0,0,'',4.35,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(366,14145,'rosinyol deshidratat','rosinyol deshidratat','','','','','','',465,0,17,1,0,0,0,'',4.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(367,14146,'cama sec','cama sec','','','','','','',465,0,3,1,0,0,0,'',4.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(368,14147,'ceps deshidratados','ceps deshidratados','','','','','','',465,0,5,1,0,0,0,'',5.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(369,14148,'trompeta deshidratada','trompeta deshidratada','','','','','','',465,0,24,1,0,0,0,'',4.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(370,14149,'mirabolat roja','mirabolat roja','','','','','','',3,0,36,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1400577445820.jpg','','',0,0.000,0),(371,14150,'cebolla caramelizada','cebolla caramelizada','','','','','','',467,0,15,1,0,0,0,'',2.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(372,14151,'manzana caramelizada','manzana caramelizada','','','','','','',467,0,35,1,0,0,0,'',2.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(373,14152,'albaricoque moniqui','albaricoque moniqui','','','','','','',3,0,1,1,0,0,0,'',3.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_053.png','','',0,0.000,0),(374,14153,'melocoton Teruel','melocoton Teruel','','','','','','',2,0,17,1,0,0,0,'',3.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_076.png','','',0,0.000,0),(375,14154,'sucar granel','sucar granel','','','','','','',6,0,48,1,0,0,0,'',1.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'TOMATE COLGAR GRANEL_1.JPG','','',0,0.000,0),(376,14155,'melocoton bandeja','melocoton bandeja','','','','','','',466,0,16,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(377,14156,'vasito fruta','vasito fruta','','','','','','',466,0,31,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(378,14157,'calanda pequeño','calanda pequeño','','','','','','',2,0,3,1,0,0,0,'',1.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_080.png','','',0,0.000,0),(379,14158,'zumo litro','zumo litro','','','','','','',462,0,28,1,0,0,0,'',1.00,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(380,14159,'pepino español','pepino español','','','','','','',6,0,37,1,0,0,0,'',2.40,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_075_1.png','','',0,0.000,0),(381,14161,'corazones alcachofa','corazones alcachofa','','','','','','',467,0,17,1,0,0,0,'',2.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(382,14162,'cebolla roja','cebolla roja','','','','','','',6,0,13,1,0,0,0,'',1.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'13795.png','','',0,0.000,0),(383,14163,'aceite germanor 750ml','aceite germanor 750ml','','','','','','',467,0,6,1,0,0,0,'',7.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(384,14164,'Col KALE','Col KALE','','','','','','',9,0,20,1,0,0,0,'',3.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_018.png','','',0,0.000,0),(385,14165,'judia fecciosa','judia fecciosa','','','','','','',9,0,31,1,0,0,0,'',4.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(386,14166,'kaqui granel','kaqui granel','','','','','','',2,0,11,1,0,0,0,'',2.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_011.png','','',0,0.000,0),(387,14167,'nuez granel','nuez granel','','','','','','',5,0,17,1,0,0,0,'',6.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'FS_019.png','','',0,0.000,0),(388,14168,'aceite 1/2l','aceite 1/2l','','','','','','',467,0,4,1,0,0,0,'',5.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(389,14169,'mix ensaladas','mix ensaladas','','','','','','',5,0,14,1,0,0,0,'',2.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'13779.png','','',0,0.000,0),(390,14170,'aceite Neus especial','aceite Neus especial','','','','','','',467,0,7,1,0,0,0,'',12.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(391,14171,'apio bandeja','apio bandeja','','','','','','',9,0,6,1,0,0,0,'',2.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13789_1.png','','',0,0.000,0),(392,14172,'zumo detox 250ml','zumo detox 250ml','','','','','','',462,0,25,1,0,0,0,'',2.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(393,14173,'detox 500','detox 500','','','','','','',462,0,9,1,0,0,0,'',5.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(394,14174,'romesco pequeña','romesco pequeña','','','','','','',467,0,50,1,0,0,0,'',2.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(395,14175,'aceite ecologico','aceite ecologico','','','','','','',467,0,5,1,0,0,0,'',6.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(396,14176,'maiz mini','maiz mini','','','','','','',6,0,26,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13838.png','','',0,0.000,0),(397,14177,'caña de azucar','caña de azucar','','','','','','',4,0,6,1,0,0,0,'',2.40,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'Cañaazúcar.jpg','','',0,0.000,0),(398,14178,'tirabeque bandeja','tirabeque bandeja','','','','','','',466,0,27,1,0,0,0,'',2.25,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_101_2.png','','',0,0.000,0),(399,14179,'alcachofa pelada','alcachofa pelada','','','','','','',466,0,2,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(400,14180,'6 huevos ecologicos','6 huevos ecologicos','','','','','','',467,0,1,1,0,0,0,'',2.75,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1379430312200_1.jpg','','',0,0.000,0),(401,14181,'mini mandarina','mini mandarina','','','','','','',1,0,8,1,0,0,0,'',5.50,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_156.png','','',0,0.000,0),(402,14182,'agua de coco','agua de coco','','','','','','',462,0,3,1,0,0,0,'',2.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(403,14183,'corazones alcachofa','corazones alcachofa','','','','','','',468,0,1,1,0,0,0,'',3.65,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(404,14184,'habitas','habitas','','','','','','',468,0,5,1,0,0,0,'',3.80,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(405,14185,'menestra verduras','menestra verduras','','','','','','',468,0,8,1,0,0,0,'',2.90,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(406,14186,'guisantes cocidos','guisantes cocidos','','','','','','',468,0,3,1,0,0,0,'',2.75,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(407,14187,'guisantes con patata','guisantes con patata','','','','','','',468,0,4,1,0,0,0,'',2.90,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(408,14188,'ensaladilla','ensaladilla','','','','','','',468,0,2,1,0,0,0,'',2.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(409,14189,'preparado verduras','preparado verduras','','','','','','',466,0,24,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(410,14190,'freson Arenys de munt','freson Arenys de munt','','','','','','',3,0,22,1,0,0,0,'',7.95,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'10011.jpg','','',0,0.000,0),(411,14191,'esparrago tucano','esparrago tucano','','','','','','',9,0,25,1,0,0,0,'',2.60,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_044_3.png','','',0,0.000,0),(412,14192,'gazpacho pequeño','gazpacho pequeño','','','','','','',462,0,14,1,0,0,0,'',2.25,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(413,14193,'maiz fresco','maiz fresco','','','','','','',6,0,25,1,0,0,0,'',2.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'MAIZ DULCE.jpg','','',0,0.000,0),(414,14194,'patata bolsa 3kg','patata bolsa 3kg','','','','','','',6,0,30,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(415,14195,'patata bolsa','patata bolsa','','','','','','',6,0,29,1,0,0,0,'',2.75,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(416,14196,'esparrago blanco','esparrago blanco','','','','','','',9,0,22,1,0,0,0,'',12.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_043.png','','',0,0.000,0),(417,14197,'maiz cocido','maiz cocido','','','','','','',468,0,7,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(418,14199,'remolacha cocida','remolacha cocida','','','','','','',468,0,10,1,0,0,0,'',1.00,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(419,14200,'gazpacho bio','gazpacho bio','','','','','','',462,0,13,1,0,0,0,'',3.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'1378752320752_1.jpg','','',0,0.000,0),(420,14201,'germinado rabano','germinado rabano','','','','','','',464,0,13,1,0,0,0,'',1.90,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(421,14202,'germinado cebolla','germinado cebolla','','','','','','',464,0,12,1,0,0,0,'',2.90,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(422,14203,'cereza caja 1kg','cereza caja 1kg','','','','','','',3,0,10,1,0,0,0,'',5.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'13748.png','','',0,0.000,0),(423,14204,'murgula','murgula','','','','','','',465,0,14,1,0,0,0,'',55.00,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(424,14205,'judia con patata','judia con patata','','','','','','',468,0,6,1,0,0,0,'',2.90,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(425,14206,'perita castells granel','perita castells granel','','','','','','',3,0,41,1,0,0,0,'',4.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_124_1.png','','',0,0.000,0),(426,14207,'tomate rebelion','tomate rebelion','','','','','','',466,0,30,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1381340302835_3.jpg','','',0,0.000,0),(427,14208,'tomate pimiento bandeja','tomate pimiento bandeja','','','','','','',466,0,28,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'1381340299118_1.jpg','','',0,0.000,0),(428,14209,'Navarra puntas','Navarra puntas','','','','','','',467,0,47,1,0,0,0,'',4.95,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(429,14210,'Navarra grande','Navarra grande','','','','','','',467,0,46,1,0,0,0,'',8.50,0,0.00,0,0,1,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(430,14211,'acelga roja','acelga roja','','','','','','',469,0,1,1,0,0,0,'',2.20,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(431,14212,'piparras','piparras','','','','','','',469,0,4,1,0,0,0,'',4.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(432,14213,'pimiento Gernika','pimiento Gernika','','','','','','',469,0,3,1,0,0,0,'',5.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(433,14214,'cebolleta roja','cebolleta roja','','','','','','',469,0,2,1,0,0,0,'',1.75,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(434,14215,'lechuga vasca','lechuga vasca','','','','','','',469,0,19,1,0,0,0,'',1.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_042.png','','',0,0.000,0),(435,14216,'Bimi','Bimi','','','','','','',9,0,7,1,0,0,0,'',1.95,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_015.png','','',0,0.000,0),(436,14217,'apio-nabo celeri','apio-nabo celeri','','','','','','',6,0,3,1,0,0,0,'',1.99,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(437,14219,'tofona estiu','tofona estiu','','','','','','',465,0,22,1,0,0,0,'',4.50,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(438,14220,'patata y cebolla','patata y cebolla','','','','','','',468,0,9,1,0,0,0,'',2.15,0,0.00,0,0,1,2,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(439,14221,'patata kennebec','patata kennebec','','','','','','',6,0,32,1,0,0,0,'',1.10,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'PATATA ROJA_1.jpg','','',0,0.000,0),(440,14222,'aguacate dominicano','aguacate dominicano','','','','','','',4,0,4,1,0,0,0,'',5.90,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'F_001.png','','',0,0.000,0),(441,14223,'yuca','yuca','','','','','','',6,0,62,1,0,0,0,'',2.60,0,0.00,1,0,1,2,'',0,0,0.000,'',0,'','-',0,'V_114.png','','',0,0.000,0);
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `itemtextbonus`
--

LOCK TABLES `itemtextbonus` WRITE;
/*!40000 ALTER TABLE `itemtextbonus` DISABLE KEYS */;
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
) ENGINE=InnoDB AUTO_INCREMENT=1226 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localcfgscale`
--

LOCK TABLES `localcfgscale` WRITE;
/*!40000 ALTER TABLE `localcfgscale` DISABLE KEYS */;
INSERT INTO `localcfgscale` VALUES (8,'RESOURCES/user-icons-path','../share/pcscale/resources/icons/item'),(12,'TOUCH/loader','1'),(14,'TOUCH/scale-icon','scale'),(15,'TOUCH/slave-icon','slave'),(108,'TOUCH/items-dnd','1'),(796,'NETWORK/scale-number','1'),(797,'CORE/mup-group','1'),(798,'CORE/plu-board-group','1'),(799,'NETWORK/scale-counter','1'),(1157,'sw-full-version','2.10 - Methone. Build id: 20160621 - 2222'),(1159,'TOUCH/vendor-icon','vnd1'),(1162,'serial-number','862052'),(1164,'TOUCH/name','Pantalla-1024x768-20Fam_Shortname-mail-800x480'),(1172,'local-ticket/StartNumber','0'),(1173,'local-ticket/MaximumNumber','0'),(1174,'TOUCH/GROUPS/scroll-ItemsFam/default','board1'),(1176,'TOUCH/GROUPS/menu/default','menu01'),(1184,'TOUCH/GROUPS/VendorsGroupFam/default','SalesFamVendor2'),(1185,'PRINTER/lbl-format','1'),(1186,'PRINTER/prn-format','estandar+logo2'),(1195,'TOUCH/GROUPS/información de articulos/default','Inf-art-sel'),(1208,'TOUCH/GROUPS/scroll-sales/default','board3'),(1211,'CORE/key-groups-linked','0'),(1215,'simplified-invoice-serial','L1600101');
/*!40000 ALTER TABLE `localcfgscale` ENABLE KEYS */;
UNLOCK TABLES;

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
) ENGINE=InnoDB AUTO_INCREMENT=249 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localcountries`
--

LOCK TABLES `localcountries` WRITE;
/*!40000 ALTER TABLE `localcountries` DISABLE KEYS */;
INSERT INTO `localcountries` VALUES (1,4,'AF','Afganistán'),(2,248,'AX','Islas Äland'),(3,8,'AL','Albania'),(4,12,'DZ','Algeria'),(5,16,'AS','Samoa Americana'),(6,20,'AD','Andorra'),(7,24,'AO','Angola'),(8,660,'AI','Anguila'),(9,10,'AQ','Antártida'),(10,28,'AG','Antigua y Barbuda'),(11,32,'AR','Argentina'),(12,51,'AM','Armenia'),(13,533,'AW','Aruba'),(14,36,'AU','Australia'),(15,40,'AT','Austria'),(16,31,'AZ','Azerbayán'),(17,44,'BS','Bahamas'),(18,48,'BH','Bahrein'),(19,50,'BD','Bangladesh'),(20,52,'BB','Barbados'),(21,112,'BY','Bielorrusia'),(22,56,'BE','Bélgica'),(23,84,'BZ','Belice'),(24,204,'BJ','Benín'),(25,60,'BM','Islas Bermudas'),(26,64,'BT','Bhután'),(27,68,'BO','Bolivia, Estado plurinacional de'),(28,535,'BQ','Bonaire, Saint Eustatius and Saba'),(29,70,'BA','Bosnia y Herzegovina'),(30,72,'BW','Botsuana'),(31,74,'BV','Isla Bouvet'),(32,76,'BR','Brasil'),(33,86,'IO','Británico del Océano Índico, Territorio'),(34,96,'BN','Brunei Darussalam'),(35,100,'BG','Bulgaria'),(36,854,'BF','Burquina Faso'),(37,108,'BI','Burundi'),(38,116,'KH','Camboya'),(39,120,'CM','Camerún'),(40,124,'CA','Canadá'),(41,132,'CV','Cabo Verde'),(42,136,'KY','Islas Caimán'),(43,140,'CF','Centro-africana, República'),(44,148,'TD','Chad'),(45,152,'CL','Chile'),(46,156,'CN','China'),(47,162,'CX','Isla de Navidad'),(48,166,'CC','Islas Cocos (Keeling)'),(49,170,'CO','Colombia'),(50,174,'KM','Comores, Islas'),(51,178,'CG','Congo'),(52,180,'CD','Congo, República Democrática del'),(53,184,'CK','Islas Cook'),(54,188,'CR','Costa Rica'),(55,384,'CI','Costa de Marfíl'),(56,191,'HR','Croacia'),(57,192,'CU','Cuba'),(58,531,'CW','Curasao'),(59,196,'CY','Chipre'),(60,203,'CZ','República Checa'),(61,208,'DK','Dinamarca'),(62,262,'DJ','Yibuti'),(63,212,'DM','Dominica'),(64,214,'DO','República Dominicana'),(65,218,'EC','Ecuador'),(66,818,'EG','Egipto'),(67,222,'SV','El Salvador'),(68,226,'GQ','Guinea Ecuatorial'),(69,232,'ER','Eritrea'),(70,233,'EE','Estonia'),(71,231,'ET','Etiopía'),(72,238,'FK','Islas Falkland (Malvinas)'),(73,234,'FO','Islas Feroe'),(74,242,'FJ','Fiyi'),(75,246,'FI','Finlandia'),(76,250,'FR','Francia'),(77,254,'GF','Guayana Francesa'),(78,258,'PF','Polinesia Francesa'),(79,260,'TF','Territorios Franceses del Sur'),(80,266,'GA','Gabón'),(81,270,'GM','Gambia'),(82,268,'GE','Georgia'),(83,276,'DE','Alemania'),(84,288,'GH','Ghana'),(85,292,'GI','Gibraltar'),(86,300,'GR','Grecia'),(87,304,'GL','Groenlandia'),(88,308,'GD','Granada'),(89,312,'GP','Guadalupe'),(90,316,'GU','Guam'),(91,320,'GT','Guatemala'),(92,831,'GG','Guernsey'),(93,324,'GN','Guinea'),(94,624,'GW','Guinea-Bissau'),(95,328,'GY','Guyana'),(96,332,'HT','Haití'),(97,334,'HM','Islas Heard, Islas y McDonald'),(98,336,'VA','Santa Sede (Ciudad Estado del Vaticano)'),(99,340,'HN','Honduras'),(100,344,'HK','Hong Kong'),(101,348,'HU','Hungría'),(102,352,'IS','Islandia'),(103,356,'IN','India'),(104,360,'ID','Indonesia'),(105,364,'IR','Irán, República islámica de'),(106,368,'IQ','Irak'),(107,372,'IE','Irlanda'),(108,833,'IM','Isla de Man'),(109,376,'IL','Israel'),(110,380,'IT','Italia'),(111,388,'JM','Jamaica'),(112,392,'JP','Japón'),(113,832,'JE','Jersey'),(114,400,'JO','Jordania'),(115,398,'KZ','Kazajistán'),(116,404,'KE','Kenia'),(117,296,'KI','Kiribati'),(118,408,'KP','Corea, República Democrática Popular de'),(119,410,'KR','Corea, República de'),(120,414,'KW','Kuwait'),(121,417,'KG','Kirgizstán'),(122,418,'LA','República Democrática Popular de Lao'),(123,428,'LV','Letonia'),(124,422,'LB','Líbano'),(125,426,'LS','Lesoto'),(126,430,'LR','Liberia'),(127,434,'LY','Libyan Arab Jamahiriya'),(128,438,'LI','Liechtenstein'),(129,440,'LT','Lituania'),(130,442,'LU','Luxemburgo'),(131,446,'MO','Macao'),(132,807,'MK','Macedonia, República de'),(133,450,'MG','Madagascar'),(134,454,'MW','Malawi'),(135,458,'MY','Malasia'),(136,462,'MV','Islas Maldivas'),(137,466,'ML','Mali'),(138,470,'MT','Malta'),(139,584,'MH','Islas Marshall'),(140,474,'MQ','Martinica'),(141,478,'MR','Mauritania'),(142,480,'MU','Mauricio'),(143,175,'YT','Mayotte'),(144,484,'MX','México'),(145,583,'FM','Micronesia, Estados Federados de'),(146,498,'MD','Moldavia, República de'),(147,492,'MC','Mónaco'),(148,496,'MN','Mongolia'),(149,499,'ME','Montenegro'),(150,500,'MS','Montserrat'),(151,504,'MA','Marruecos'),(152,508,'MZ','Mozambique'),(153,104,'MM','Birmania'),(154,516,'NA','Namibia'),(155,520,'NR','Nauru'),(156,524,'NP','Nepal'),(157,528,'NL','Países Bajos'),(158,540,'NC','Nueva Caledonia'),(159,554,'NZ','Nueva Zelanda'),(160,558,'NI','Nicaragua'),(161,562,'NE','Niger'),(162,566,'NG','Nigeria'),(163,570,'NU','Niue'),(164,574,'NF','Isla Norfolk'),(165,580,'MP','Islas Marianas del Norte'),(166,578,'NO','Noruega'),(167,512,'OM','Omán'),(168,586,'PK','Pakistán'),(169,585,'PW','Palau'),(170,275,'PS','Territorio palestino, Ocupado'),(171,591,'PA','Panamá'),(172,598,'PG','Papúa Nueva Guinea'),(173,600,'PY','Paraguay'),(174,604,'PE','Perú'),(175,608,'PH','Filipinas'),(176,612,'PN','Pitcairn'),(177,616,'PL','Polonia'),(178,620,'PT','Portugal'),(179,630,'PR','Puerto Rico'),(180,634,'QA','Qatar'),(181,638,'RE','Reunión'),(182,642,'RO','Rumanía'),(183,643,'RU','Federación Rusa'),(184,646,'RW','Ruanda'),(185,652,'BL','San Bartolomé'),(186,654,'SH','Santa Elena, Ascensión y Tristán de Acuña'),(187,659,'KN','San Cristobo y Nevis'),(188,662,'LC','Santa Lucía'),(189,663,'MF','San Martín (zona francesa)'),(190,666,'PM','San Pedro y Miquelon'),(191,670,'VC','San Vicente y las Granadinas'),(192,882,'WS','Samoa'),(193,674,'SM','San Marino'),(194,678,'ST','Santo Tomé y Príncipe'),(195,682,'SA','Arabia Saudí'),(196,686,'SN','Senegal'),(197,688,'RS','Serbia'),(198,690,'SC','Seychelles'),(199,694,'SL','Sierra Leona'),(200,702,'SG','Singapur'),(201,534,'SX','Sint Maarten'),(202,703,'SK','Eslovaquia'),(203,705,'SI','Eslovenia'),(204,90,'SB','Islas Salomón'),(205,706,'SO','Somalia'),(206,710,'ZA','Suráfrica'),(207,239,'GS','Georgia del Sur e Islas Sandwitch del Sur'),(208,724,'ES','España'),(209,144,'LK','Sri Lanka'),(210,736,'SD','Sudán'),(211,740,'SR','Surinám'),(212,744,'SJ','Svalbard y Jan Mayen'),(213,748,'SZ','Swazilandia'),(214,752,'SE','Suecia'),(215,756,'CH','Suiza'),(216,760,'SY','República árabe de Siria'),(217,158,'TW','Taiwán, Provincia de China'),(218,762,'TJ','Tadjikistán'),(219,834,'TZ','Tanzania, República unida de'),(220,764,'TH','Tailandia'),(221,626,'TL','Timor-Leste'),(222,768,'TG','Togo'),(223,772,'TK','Tokelau'),(224,776,'TO','Tonga'),(225,780,'TT','Trinidad y Tobago'),(226,788,'TN','Tunez'),(227,792,'TR','Turquía'),(228,795,'TM','Turkmenistán'),(229,796,'TC','Turks y Caicos, Islas'),(230,798,'TV','Tuvalu'),(231,800,'UG','Uganda'),(232,804,'UA','Ucrania'),(233,784,'AE','Emiratos Árabes Unidos'),(234,826,'GB','Reino Unido'),(235,840,'US','Estados Unidos'),(236,581,'UM','Islas menores exteriores de Estados Unidos'),(237,858,'UY','Uruguay'),(238,860,'UZ','Uzbekistán'),(239,548,'VU','Vanuatu'),(240,862,'VE','Venezuela, Bolivarian republic of'),(241,704,'VN','Vietnam'),(242,92,'VG','Islas Vírgenes, Británicas'),(243,850,'VI','Islas Vírgenes, de EEUU'),(244,876,'WF','Wallis y Futuna'),(245,732,'EH','Sahara Occidental'),(246,887,'YE','Yemen'),(247,894,'ZM','Zambia'),(248,716,'ZW','Zimbabue');
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localdatalogger`
--

LOCK TABLES `localdatalogger` WRITE;
/*!40000 ALTER TABLE `localdatalogger` DISABLE KEYS */;
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
) ENGINE=InnoDB AUTO_INCREMENT=21 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localerasetickets`
--

LOCK TABLES `localerasetickets` WRITE;
/*!40000 ALTER TABLE `localerasetickets` DISABLE KEYS */;
INSERT INTO `localerasetickets` VALUES (18,'T','2016-10-28 10:16:29'),(19,'P','2016-10-28 10:07:44'),(20,'S','2016-10-28 10:07:40');
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
  `SectionItem` int(11) DEFAULT '0',
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
  `StartDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `TimesPrinted` tinyint(3) unsigned NOT NULL DEFAULT '0',
  `DatTim` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadFrom` int(11) DEFAULT '0',
  `ReloadScale` int(11) NOT NULL DEFAULT '0',
  `ReloadDateTime` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `ReloadTicketNumber` int(11) NOT NULL DEFAULT '0',
  `ReloadtckType` char(1) DEFAULT '',
  `Customer` varchar(16) NOT NULL DEFAULT '',
  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
  `Invoice` varchar(20) NOT NULL DEFAULT '',
  `InvoiceDate` date DEFAULT '0000-00-00',
  `tckDiscount` decimal(4,2) DEFAULT '0.00',
  `ParkedNumber` int(11) DEFAULT '0',
  `NeedCopy` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`Id`),
  KEY `Vendor` (`Vendor`),
  KEY `IdNetHticket` (`IdNetHticket`),
  KEY `Invoice` (`Invoice`),
  KEY `datim_idx` (`DatTim`)
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
    IF(OLD.`tckDiscount`!=NEW.`tckDiscount`) THEN
      SET @frac_digits := IFNULL(GLOBAL_GET("frac_digits"),2);
      SET NEW.`NetAmount`=ROUND(NEW.`GrossAmount`*((100-NEW.`tckDiscount`)/100), @frac_digits);
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
  `Weight` decimal(5,3) DEFAULT NULL,
  `Tare` decimal(5,3) DEFAULT '0.000',
  `TTare` char(1) DEFAULT '',
  `lineDiscount` decimal(4,2) DEFAULT NULL,
  `Price` decimal(8,2) DEFAULT NULL,
  `Amount` decimal(10,2) DEFAULT NULL,
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
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalLTicketsInserted` AFTER INSERT ON `LocalLTickets`
FOR EACH ROW BEGIN
  SET @frac_digits := IFNULL(GLOBAL_GET("frac_digits"),2);

  IF(NEW.`NetStat`='L') THEN

    IF(NEW.`ReloadScale`=0) THEN
      SET @wgt = IF(NEW.`SaleForm`=1,NEW.`Weight`,0);
      SET @units = IF(NEW.`SaleForm`=0,NEW.`Weight`,0);
      UPDATE `LocalHTickets` SET `NumLines`=`NumLines`+1, `AddedLines`=`AddedLines`+1, `GrossAmount`=`GrossAmount`+NEW.`Amount`,
             `NetAmount`=ROUND(`GrossAmount`*((100-`tckDiscount`)/100), @frac_digits),
             `NegativeLines`=`NegativeLines`+IF(NEW.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`+IF(NEW.`Price`<0,-NEW.`Amount`,0),
             `TotalWeight`=`TotalWeight`+IF(NEW.`Price`<0,-@wgt,@wgt), `TotalUnits`=`TotalUnits`+IF(NEW.`Price`<0,-@units,@units),
             `DatTim`=NOW(), `SectionItem`=NEW.`Section` WHERE `LocalHTickets`.`Id`=NEW.`IdHTicket` AND `LocalHTickets`.`NetStat`='L';
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
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LocalLTicketsUpdated` AFTER UPDATE ON `LocalLTickets`
FOR EACH ROW BEGIN
  SET @frac_digits := IFNULL(GLOBAL_GET("frac_digits"),2);
  IF(NEW.`lineType`!=OLD.`lineType`) THEN
    SET @wgt = IF(OLD.`SaleForm`=1,OLD.`Weight`,0);
    SET @units = IF(NEW.`SaleForm`=0,NEW.`Weight`,0);
    IF(NEW.`lineType`=0) THEN

        UPDATE `LocalHTickets` SET `AddedLines`=`AddedLines`-1, `GrossAmount`=`GrossAmount`-OLD.`Amount`,
               `NetAmount`=ROUND(`GrossAmount`*((100-`tckDiscount`)/100), @frac_digits),
               `NegativeLines`=`NegativeLines`-IF(OLD.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`-IF(OLD.`Price`<0,-OLD.`Amount`,0),
               `CanceledLines`=`CanceledLines`+1, `CanceledAmount` = `CanceledAmount`+OLD.`Amount`,
               `TotalWeight`=`TotalWeight`-IF(OLD.`Price`<0,-@wgt,@wgt), `TotalUnits`=`TotalUnits`-IF(NEW.`Price`<0,-@units,@units),
               `DatTim`=NOW() WHERE `LocalHTickets`.`Id`=OLD.`IdHTicket`;
    ELSE
        UPDATE `LocalHTickets` SET `AddedLines`=`AddedLines`+1, `GrossAmount`=`GrossAmount`+OLD.`Amount`,
               `NetAmount`=ROUND(`GrossAmount`*((100-`tckDiscount`)/100), @frac_digits),
               `NegativeLines`=`NegativeLines`+IF(OLD.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`+IF(OLD.`Price`<0,-OLD.`Amount`,0),
               `CanceledLines`=`CanceledLines`-1, `CanceledAmount` = `CanceledAmount`-OLD.`Amount`,
               `TotalWeight`=`TotalWeight`+IF(OLD.`Price`<0,-@wgt,@wgt), `TotalUnits`=`TotalUnits`+IF(NEW.`Price`<0,-@units,@units),
               `DatTim`=NOW() WHERE `LocalHTickets`.`Id`=OLD.`IdHTicket`;
    END IF;
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
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localnumtickets`
--

LOCK TABLES `localnumtickets` WRITE;
/*!40000 ALTER TABLE `localnumtickets` DISABLE KEYS */;
INSERT INTO `localnumtickets` VALUES (2,'I','L1600102',1);
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
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `lineType` smallint(6) DEFAULT NULL,
  `Item` int(11) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `Weight` decimal(5,3) DEFAULT NULL,
  `Tare` decimal(5,3) DEFAULT '0.000',
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
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
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `lineType` smallint(6) DEFAULT NULL,
  `Item` int(11) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `Weight` decimal(5,3) DEFAULT NULL,
  `Tare` decimal(5,3) DEFAULT '0.000',
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
  `CheckFamilies` tinyint(1) DEFAULT '0',
  `CheckCounterFamilies` tinyint(1) DEFAULT '0',
  `CheckItemGroupKeysText` tinyint(1) DEFAULT '0',
  `CheckTareCodes` tinyint(1) DEFAULT '0',
  `CheckScripts` tinyint(1) DEFAULT '0',
  `LocalTickets` tinyint(1) DEFAULT '0',
  `AdvertisingChanged` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localstatus`
--

LOCK TABLES `localstatus` WRITE;
/*!40000 ALTER TABLE `localstatus` DISABLE KEYS */;
INSERT INTO `localstatus` VALUES (1,71,'2.10.2222','es_ES',0,0,0,0,0,0,0,0,0,0);
/*!40000 ALTER TABLE `localstatus` ENABLE KEYS */;
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
INSERT INTO `localvendorkeys` VALUES (1,1,101),(1,2,102),(1,3,103),(1,4,104),(1,5,105),(1,6,106),(1,7,107),(1,8,108),(1,9,109),(1,10,110);
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
      UPDATE `LocalVendorRegister` SET `ReleaseDateTime`=NOW() WHERE `Vendor`=OLD.`Vendor` AND `ReleaseDateTime`=0;
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
  `Tare` decimal(5,3) DEFAULT NULL,
  `Discount` decimal(4,2) DEFAULT NULL,
  `Lot` varchar(20) NOT NULL DEFAULT '',
  `Ordinal` int(10) unsigned DEFAULT NULL,
  `ExpiringDate` date NOT NULL DEFAULT '0000-00-00',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Vendor` (`Vendor`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localvendorpacket`
--

LOCK TABLES `localvendorpacket` WRITE;
/*!40000 ALTER TABLE `localvendorpacket` DISABLE KEYS */;
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
) ENGINE=InnoDB AUTO_INCREMENT=1776 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localworkingmodes`
--

LOCK TABLES `localworkingmodes` WRITE;
/*!40000 ALTER TABLE `localworkingmodes` DISABLE KEYS */;
INSERT INTO `localworkingmodes` VALUES (1,'','config/CORE/USER-CHECK-MENU-FORM','ON'),(6,'00 sales','aftersale','0'),(13,'00 sales','based','sales'),(34,'00 sales','time-aftersale','1'),(46,'01 packet','based','packet'),(49,'01 packet','description','Envasado'),(50,'01 packet','in-use','0'),(52,'01 packet','keyboard','packet'),(53,'01 packet','keyboard-loop','packet-loop'),(60,'01 packet','return-weight','ON'),(64,'01 packet','user','admin'),(67,'02 selfservice','amount-zero','OFF'),(68,'02 selfservice','autocode','OFF'),(69,'02 selfservice','based','selfservice'),(70,'02 selfservice','change-price','OFF'),(71,'02 selfservice','code-digits','6'),(72,'02 selfservice','constant-key','ON'),(75,'02 selfservice','item-access','1'),(77,'02 selfservice','negative-amount','OFF'),(78,'02 selfservice','pluprc-key','OFF'),(79,'02 selfservice','return-weight','ON'),(80,'02 selfservice','sections',''),(81,'02 selfservice','show-print-ss','ON'),(82,'02 selfservice','time-error-ss','3'),(84,'02 selfservice','user-check-mode','OFF'),(85,'02 selfservice','user-check-mode-form','ON'),(93,'03 pos','based','sales'),(310,'00 sales','allow-reload','ON'),(483,'01 packet','user-check-mode-form','OFF'),(484,'01 packet','user-check-mode','OFF'),(485,'01 packet','user-prog-prices','ON'),(486,'01 packet','sections',''),(487,'01 packet','packet-customer','ON'),(488,'01 packet','show-print-packet','ON'),(489,'01 packet','allow-mult-key','ON'),(490,'01 packet','packet-vendor','0'),(491,'01 packet','show-form-packet','0'),(492,'01 packet','packet-after-item','OFF'),(493,'01 packet','autocode','ON'),(494,'01 packet','pluprc-key','ON'),(495,'01 packet','constant-key','ON'),(496,'01 packet','change-price','ON'),(497,'01 packet','amount-zero','ON'),(498,'01 packet','negative-amount','ON'),(499,'01 packet','code-digits','6'),(500,'01 packet','item-access','0'),(501,'01 packet','mixed-allowed','OFF'),(502,'01 packet','weightlow-tsec','0'),(503,'01 packet','weightlow-beep','ON'),(504,'01 packet','weightlow-zero','OFF'),(619,'03 pos','allow-reload','ON'),(1006,'03 pos','description','POS'),(1007,'03 pos','keyboard','POSFamilies'),(1008,'03 pos','user','admin'),(1009,'03 pos','key-fnc-on-pressed','0'),(1010,'03 pos','keyboard-loop',''),(1011,'03 pos','keyboard-vendors','POSVendor'),(1012,'03 pos','PLUload-func','2'),(1013,'03 pos','PLUload-mode','1'),(1014,'03 pos','in-use','0'),(1187,'03 pos','user-check-mode-form','OFF'),(1188,'03 pos','user-check-mode','OFF'),(1189,'03 pos','user-prog-prices','ON'),(1190,'03 pos','sections',''),(1191,'03 pos','print-ticket','ON'),(1192,'03 pos','print-label-total','OFF'),(1193,'03 pos','print-label-line','OFF'),(1194,'03 pos','close-with-line','OFF'),(1195,'03 pos','total-discount','ON'),(1196,'03 pos','sales-customer','ON'),(1197,'03 pos','allow-orders','ON'),(1198,'03 pos','payment','2'),(1199,'03 pos','startlocal','0'),(1200,'03 pos','allow-cancel','ON'),(1201,'03 pos','allow-amount0','ON'),(1202,'03 pos','auto-fix','ON'),(1203,'03 pos','times-print','250'),(1205,'03 pos','auto-vendor','ON'),(1206,'03 pos','customer-care','OFF'),(1207,'03 pos','park-ticket','ON'),(1208,'03 pos','ticket-format',''),(1209,'03 pos','label-format',''),(1210,'03 pos','parking-format','parked'),(1211,'03 pos','reload-type','48'),(1212,'03 pos','reprint-type','240'),(1213,'03 pos','allow-training','OFF'),(1214,'03 pos','autocode','OFF'),(1215,'03 pos','pluprc-key','ON'),(1216,'03 pos','constant-key','ON'),(1217,'03 pos','change-price','ON'),(1218,'03 pos','return-weight','ON'),(1219,'03 pos','amount-zero','ON'),(1220,'03 pos','negative-amount','ON'),(1221,'03 pos','code-digits','6'),(1222,'03 pos','item-access','0'),(1223,'03 pos','pos-method','ON'),(1224,'03 pos','clear-tare','ON'),(1225,'03 pos','weightlow-tsec','0'),(1226,'03 pos','weightlow-beep','ON'),(1227,'03 pos','weightlow-zero','OFF'),(1228,'','user-check-menu-form','OFF'),(1291,'00 sales','description','Venta'),(1292,'00 sales','keyboard','sales'),(1293,'00 sales','user','admin'),(1294,'00 sales','key-fnc-on-pressed','0'),(1295,'00 sales','keyboard-loop',''),(1296,'00 sales','keyboard-vendors',''),(1297,'00 sales','PLUload-func','2'),(1298,'00 sales','PLUload-mode','1'),(1299,'00 sales','in-use','1'),(1302,'00 sales','user-check-mode-form','OFF'),(1303,'00 sales','user-check-mode','OFF'),(1304,'00 sales','user-prog-prices','ON'),(1305,'00 sales','sections',''),(1306,'00 sales','print-ticket','ON'),(1307,'00 sales','print-label-total','OFF'),(1308,'00 sales','print-label-line','OFF'),(1309,'00 sales','close-with-line','OFF'),(1310,'00 sales','total-discount','ON'),(1311,'00 sales','sales-customer','ON'),(1312,'00 sales','allow-orders','OFF'),(1313,'00 sales','payment','2'),(1314,'00 sales','startlocal','0'),(1315,'00 sales','allow-cancel','ON'),(1316,'00 sales','allow-amount0','ON'),(1317,'00 sales','auto-fix','OFF'),(1318,'00 sales','times-print','250'),(1320,'00 sales','auto-vendor','ON'),(1321,'00 sales','customer-care','OFF'),(1322,'00 sales','park-ticket','ON'),(1323,'00 sales','ticket-format',''),(1324,'00 sales','label-format',''),(1325,'00 sales','parking-format','parked'),(1326,'00 sales','reload-type','48'),(1327,'00 sales','reprint-type','240'),(1328,'00 sales','allow-training','OFF'),(1329,'00 sales','autocode','ON'),(1330,'00 sales','pluprc-key','ON'),(1331,'00 sales','constant-key','ON'),(1332,'00 sales','change-price','ON'),(1333,'00 sales','return-weight','ON'),(1334,'00 sales','amount-zero','ON'),(1335,'00 sales','negative-amount','ON'),(1336,'00 sales','code-digits','6'),(1337,'00 sales','item-access','0'),(1338,'00 sales','pos-method','OFF'),(1339,'00 sales','clear-tare','ON'),(1340,'00 sales','weightlow-tsec','0'),(1341,'00 sales','weightlow-beep','ON'),(1342,'00 sales','weightlow-zero','OFF'),(1353,'02 selfservice','description','Autoservicio'),(1354,'02 selfservice','keyboard','selfservice'),(1355,'02 selfservice','user','admin'),(1356,'02 selfservice','key-fnc-on-pressed','1'),(1357,'02 selfservice','keyboard-loop',''),(1358,'02 selfservice','keyboard-vendors',''),(1359,'02 selfservice','PLUload-func','2'),(1360,'02 selfservice','PLUload-mode','1'),(1361,'02 selfservice','in-use','0'),(1362,'','working-mode','00 sales'),(1375,'01 packet','key-fnc-on-pressed','0'),(1377,'01 packet','keyboard-vendors',''),(1378,'01 packet','PLUload-func','2'),(1379,'01 packet','PLUload-mode','1'),(1471,'00 sales','vendor-hide-code','OFF'),(1472,'00 sales','vendor-summary','OFF'),(1473,'00 sales','single-section','OFF'),(1474,'00 sales','ean13-section-item','OFF'),(1475,'00 sales','ticket-days','1'),(1476,'00 sales','simplified-invoice','ON'),(1477,'00 sales','hold-plusec','ON'),(1478,'00 sales','plukey-dnd','Full'),(1520,'01 packet','edit-bovine-card','OFF'),(1521,'01 packet','label-format',''),(1522,'01 packet','hold-plusec','ON'),(1523,'01 packet','plukey-dnd','Full'),(1546,'03 pos','vendor-hide-code','OFF'),(1547,'03 pos','vendor-summary','OFF'),(1548,'03 pos','single-section','OFF'),(1549,'03 pos','ean13-section-item','OFF'),(1550,'03 pos','ticket-days','1'),(1551,'03 pos','simplified-invoice','OFF'),(1552,'03 pos','hold-plusec','ON'),(1553,'03 pos','plukey-dnd','Full'),(1596,'00 sales','current-data-mode','0'),(1597,'00 sales','hold-plu-to-set','Full'),(1598,'01 packet','hold-plu-to-set','Full'),(1599,'03 pos','hold-plu-to-set','Full'),(1601,'03 pos','reprint-ticket','ON'),(1602,'03 pos','reprint-label-total','OFF'),(1603,'00 sales','reprint-ticket','ON'),(1604,'00 sales','reprint-label-total','OFF'),(1660,'04 salesturn','based','sales'),(1661,'04 salesturn','description','Venta+Turno'),(1662,'04 salesturn','keyboard','salesTurn'),(1663,'04 salesturn','user','admin'),(1664,'04 salesturn','key-fnc-on-pressed','0'),(1665,'04 salesturn','PLUload-func','2'),(1666,'04 salesturn','PLUload-mode','1'),(1667,'04 salesturn','in-use','1'),(1669,'00 sales','user-pricerates','ON'),(1706,'00 sales','weight-min','0,000'),(1707,'00 sales','amount-min','0,00'),(1708,'00 sales','pricerate','0'),(1723,'04 salesturn','user-check-mode-form','OFF'),(1724,'04 salesturn','user-check-mode','OFF'),(1725,'04 salesturn','user-prog-prices','ON'),(1726,'04 salesturn','print-ticket','ON'),(1727,'04 salesturn','print-label-total','OFF'),(1728,'04 salesturn','print-label-line','OFF'),(1729,'04 salesturn','reprint-ticket','ON'),(1730,'04 salesturn','reprint-label-total','OFF'),(1731,'04 salesturn','close-with-line','OFF'),(1732,'04 salesturn','total-discount','ON'),(1733,'04 salesturn','sales-customer','ON'),(1734,'04 salesturn','allow-orders','OFF'),(1735,'04 salesturn','payment','2'),(1736,'04 salesturn','startlocal','0'),(1737,'04 salesturn','allow-cancel','ON'),(1738,'04 salesturn','allow-amount0','ON'),(1739,'04 salesturn','auto-fix','OFF'),(1740,'04 salesturn','times-print','250'),(1741,'04 salesturn','pos-method','OFF'),(1742,'04 salesturn','auto-vendor','ON'),(1743,'04 salesturn','vendor-hide-code','OFF'),(1744,'04 salesturn','vendor-summary','OFF'),(1745,'04 salesturn','customer-care','OFF'),(1746,'04 salesturn','park-ticket','ON'),(1747,'04 salesturn','single-section','OFF'),(1748,'04 salesturn','ean13-section-item','OFF'),(1749,'04 salesturn','ticket-format',''),(1750,'04 salesturn','label-format',''),(1751,'04 salesturn','parking-format','parked'),(1752,'04 salesturn','reload-type','48'),(1753,'04 salesturn','reprint-type','240'),(1754,'04 salesturn','ticket-days','1'),(1755,'04 salesturn','allow-training','OFF'),(1756,'04 salesturn','simplified-invoice','ON'),(1757,'04 salesturn','hold-plu-to-set','Full'),(1758,'04 salesturn','plukey-dnd','Full'),(1759,'04 salesturn','user-pricerates','ON'),(1760,'04 salesturn','autocode','ON'),(1761,'04 salesturn','pluprc-key','ON'),(1762,'04 salesturn','constant-key','ON'),(1763,'04 salesturn','change-price','ON'),(1764,'04 salesturn','return-weight','ON'),(1765,'04 salesturn','amount-zero','ON'),(1766,'04 salesturn','negative-amount','ON'),(1767,'04 salesturn','code-digits','6'),(1768,'04 salesturn','item-access','0'),(1769,'04 salesturn','clear-tare','ON'),(1770,'04 salesturn','weightlow-tsec','0'),(1771,'04 salesturn','weightlow-beep','ON'),(1772,'04 salesturn','weightlow-zero','OFF'),(1773,'04 salesturn','weight-min','0,000'),(1774,'04 salesturn','amount-min','0,00'),(1775,'04 salesturn','pricerate','0');
/*!40000 ALTER TABLE `localworkingmodes` ENABLE KEYS */;
UNLOCK TABLES;

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
  `Weight` decimal(5,3) DEFAULT NULL,
  `Tare` decimal(5,3) DEFAULT '0.000',
  `TTare` char(1) DEFAULT '',
  `lineDiscount` decimal(4,2) DEFAULT NULL,
  `Price` decimal(8,2) DEFAULT NULL,
  `Amount` decimal(10,2) DEFAULT NULL,
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
  PRIMARY KEY (`Id`),
  KEY `IdHTicket` (`IdHTicket`),
  CONSTRAINT `ltickets_ibfk_1` FOREIGN KEY (`IdHTicket`) REFERENCES `htickets` (`Id`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `ltickets`
--

LOCK TABLES `ltickets` WRITE;
/*!40000 ALTER TABLE `ltickets` DISABLE KEYS */;
INSERT INTO `ltickets` VALUES (1,1,'N',1,1,'-',13720,'Limas',0,1,0,1,105,'2016-10-28 10:20:56','P',0,2,4.00,2.000,0.000,'',0.00,3.99,7.98,0,'','','0000-00-00',25.00,1,'2016-10-28 10:21:13',1,'T',0,'0000-00-00 00:00:00',0,''),(2,1,'N',1,1,'-',13950,'bollo etiqueta negra',0,2,0,1,105,'2016-10-28 10:21:01','P',0,2,4.00,2.000,0.000,'',0.00,2.80,5.60,0,'','','0000-00-00',25.00,1,'2016-10-28 10:21:13',1,'T',0,'0000-00-00 00:00:00',0,''),(3,2,'N',1,1,'-',14067,'albericoc orange',0,3,0,1,105,'2016-10-28 10:21:27','P',1,2,4.00,2.000,0.000,'',0.00,1.00,2.00,0,'','','0000-00-00',0.00,1,'2016-10-28 10:21:44',2,'T',0,'0000-00-00 00:00:00',0,''),(4,2,'N',1,0,'-',14110,'azafran',0,463,0,1,105,'2016-10-28 10:21:39','P',1,1,10.00,1.000,0.000,'',0.00,1.50,1.50,0,'','','0000-00-00',0.00,1,'2016-10-28 10:21:44',2,'T',0,'0000-00-00 00:00:00',0,''),(5,3,'N',1,1,'-',13846,'albericoc mirlo',0,3,0,1,105,'2016-10-28 10:21:58','P',1,2,4.00,2.000,0.000,'',0.00,5.00,10.00,0,'','','0000-00-00',10.00,1,'2016-10-28 10:22:06',3,'T',0,'0000-00-00 00:00:00',0,'');
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
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LTicketsInserted` AFTER INSERT ON `LTickets`
FOR EACH ROW BEGIN
  SET @frac_digits := IFNULL(GLOBAL_GET("frac_digits"),2);

  IF(NEW.`NetStat`='N') THEN

    IF(NEW.`ReloadScale`=0) THEN
      SET @wgt = IF(NEW.`SaleForm`=1,NEW.`Weight`,0);
      SET @units = IF(NEW.`SaleForm`=0,NEW.`Weight`,0);
      UPDATE `HTickets` SET `NumLines`=`NumLines`+1, `AddedLines`=`AddedLines`+1, `GrossAmount`=`GrossAmount`+NEW.`Amount`,
             `NetAmount`=ROUND(`GrossAmount`*((100-`tckDiscount`)/100), @frac_digits),
             `NegativeLines`=`NegativeLines`+IF(NEW.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`+IF(NEW.`Price`<0,-NEW.`Amount`,0),
             `TotalWeight`=`TotalWeight`+IF(NEW.`Price`<0,-@wgt,@wgt), `TotalUnits`=`TotalUnits`+IF(NEW.`Price`<0,-@units,@units),
             `DatTim`=NOW(), `SectionItem`=NEW.`Section` WHERE `HTickets`.`Id`=NEW.`IdHTicket`;
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
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `LTicketsUpdated` AFTER UPDATE ON `LTickets`
FOR EACH ROW BEGIN
  SET @frac_digits := IFNULL(GLOBAL_GET("frac_digits"),2);
  IF(NEW.`lineType`!=OLD.`lineType`) THEN
    SET @wgt = IF(OLD.`SaleForm`=1,OLD.`Weight`,0);
    SET @units = IF(NEW.`SaleForm`=0,NEW.`Weight`,0);
    IF(NEW.`lineType`=0) THEN

        UPDATE `HTickets` SET `AddedLines`=`AddedLines`-1, `GrossAmount`=`GrossAmount`-OLD.`Amount`,
               `NetAmount`=ROUND(`GrossAmount`*((100-`tckDiscount`)/100), @frac_digits),
               `NegativeLines`=`NegativeLines`-IF(OLD.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`-IF(OLD.`Price`<0,-OLD.`Amount`,0),
               `CanceledLines`=`CanceledLines`+1, `CanceledAmount` = `CanceledAmount`+OLD.`Amount`,
               `TotalWeight`=`TotalWeight`-IF(OLD.`Price`<0,-@wgt,@wgt), `TotalUnits`=`TotalUnits`-IF(NEW.`Price`<0,-@units,@units),
               `DatTim`=NOW() WHERE `HTickets`.`Id`=OLD.`IdHTicket`;
    ELSE
        UPDATE `HTickets` SET `AddedLines`=`AddedLines`+1, `GrossAmount`=`GrossAmount`+OLD.`Amount`,
               `NetAmount`=ROUND(`GrossAmount`*((100-`tckDiscount`)/100), @frac_digits),
               `NegativeLines`=`NegativeLines`+IF(OLD.`Price`<0,1,0), `NegativeAmount` = `NegativeAmount`+IF(OLD.`Price`<0,-OLD.`Amount`,0),
               `CanceledLines`=`CanceledLines`-1, `CanceledAmount` = `CanceledAmount`-OLD.`Amount`,
               `TotalWeight`=`TotalWeight`+IF(OLD.`Price`<0,-@wgt,@wgt), `TotalUnits`=`TotalUnits`+IF(NEW.`Price`<0,-@units,@units),
               `DatTim`=NOW() WHERE `HTickets`.`Id`=OLD.`IdHTicket`;
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
INSERT INTO `netstatus` VALUES (1,'Pruebas','Pruebas 2016/10/28 10:17:19 3199874124',2335);
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
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `numberingtickets`
--

LOCK TABLES `numberingtickets` WRITE;
/*!40000 ALTER TABLE `numberingtickets` DISABLE KEYS */;
INSERT INTO `numberingtickets` VALUES (2,'T',1,1,1,99999,99999,1,'D',0,0,0,0),(3,'P',1,0,1,99999,9999,1,'D',0,0,0,0),(4,'S',1,0,1,99999,9999,1,'D',0,0,0,0),(5,'C',1,0,1,99999,9999,0,'D',0,0,0,0),(6,'D',1,0,1,9999,9999,0,'D',0,0,0,0),(7,'F',1,0,1,9999,9999,0,'D',0,0,0,0),(8,'Z',1,0,1,9999,9999,0,'D',0,0,0,0);
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
INSERT INTO `numinvoices` VALUES (1,99999,7,'2016/',0,'R');
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
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `numtickets`
--

LOCK TABLES `numtickets` WRITE;
/*!40000 ALTER TABLE `numtickets` DISABLE KEYS */;
INSERT INTO `numtickets` VALUES (2,'I','1600101',109),(3,'D','1',18),(4,'F','1',2),(5,'T','1',3);
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
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `lineType` smallint(6) DEFAULT NULL,
  `Item` int(11) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `Weight` decimal(5,3) DEFAULT NULL,
  `Tare` decimal(5,3) DEFAULT '0.000',
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
  `Lot` varchar(20) NOT NULL DEFAULT '',
  `Ordinal` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
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
) ENGINE=InnoDB AUTO_INCREMENT=15 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `paymethods`
--

LOCK TABLES `paymethods` WRITE;
/*!40000 ALTER TABLE `paymethods` DISABLE KEYS */;
INSERT INTO `paymethods` VALUES (8,0,'EFECTIVO',1),(9,1,'T.CREDITO',1),(10,2,'CHEQUE',0),(11,3,'PAYMETHOD3',0),(12,4,'PAYMETHOD4',0),(13,5,'PAYMETHOD5',0),(14,6,'PAYMETHOD6',0);
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
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `paytickets`
--

LOCK TABLES `paytickets` WRITE;
/*!40000 ALTER TABLE `paytickets` DISABLE KEYS */;
INSERT INTO `paytickets` VALUES (1,1,0,10.19),(2,2,0,3.50),(3,3,0,9.00);
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
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Name` (`Name`,`PlayIndex`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `playlists`
--

LOCK TABLES `playlists` WRITE;
/*!40000 ALTER TABLE `playlists` DISABLE KEYS */;
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
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
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `PLUKeysInserted` AFTER INSERT ON `PLUKeys`
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
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `PLUKeysUpdated` AFTER UPDATE ON `PLUKeys`
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
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
DELIMITER ;;
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `PLUKeysDeleted` AFTER DELETE ON `PLUKeys`
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
INSERT INTO `pricerates` VALUES (0,'');
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
/*!40000 ALTER TABLE `qlevels` ENABLE KEYS */;
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
  `Serialnumber` varchar(10) DEFAULT '',
  `System` varchar(32) DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `MAC` (`MAC`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `scales`
--

LOCK TABLES `scales` WRITE;
/*!40000 ALTER TABLE `scales` DISABLE KEYS */;
INSERT INTO `scales` VALUES (6,'00:30:18:C0:35:CF','192.168.1.211',1,1,'','S','2.10.2222','862052','Ubuntu_precise');
/*!40000 ALTER TABLE `scales` ENABLE KEYS */;
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
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
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
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES' */ ;
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
  `Name` varchar(127) NOT NULL,
  `DateTime` datetime DEFAULT NULL,
  `Weight` decimal(5,3) DEFAULT NULL,
  `Units` int(11) DEFAULT NULL,
  `Status` smallint(6) DEFAULT '0',
  PRIMARY KEY (`Id`)
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
  `Number` int(11) DEFAULT '0',
  `Counter` int(11) DEFAULT '0',
  `Scale` int(11) DEFAULT '0',
  `Vendor` int(11) DEFAULT NULL,
  `lineType` smallint(6) DEFAULT NULL,
  `Item` int(11) DEFAULT NULL,
  `SaleForm` smallint(6) DEFAULT NULL,
  `Weight` decimal(5,3) DEFAULT NULL,
  `Tare` decimal(5,3) DEFAULT '0.000',
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
  `Lot` varchar(20) NOT NULL DEFAULT '',
  `Ordinal` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
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
  `Version` smallint(6) DEFAULT '71',
  `AppVersion` varchar(16) DEFAULT '',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `status`
--

LOCK TABLES `status` WRITE;
/*!40000 ALTER TABLE `status` DISABLE KEYS */;
INSERT INTO `status` VALUES (1,71,'2.10.2222');
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
  `Tare` decimal(5,3) DEFAULT '0.000',
  `Name` varchar(20) NOT NULL DEFAULT '',
  PRIMARY KEY (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tarecodes`
--

LOCK TABLES `tarecodes` WRITE;
/*!40000 ALTER TABLE `tarecodes` DISABLE KEYS */;
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
  PRIMARY KEY (`Id`),
  UNIQUE KEY `dailycnt_idx` (`Date`,`Type`,`NetStat`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totaldaily`
--

LOCK TABLES `totaldaily` WRITE;
/*!40000 ALTER TABLE `totaldaily` DISABLE KEYS */;
INSERT INTO `totaldaily` VALUES (1,'T','N','2016-10-28',18,1,0,0,3,5,22.69,4.39,8.000,1,0,0.00,0,0.00,3,22.69);
/*!40000 ALTER TABLE `totaldaily` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `totalpaytickets`
--

DROP TABLE IF EXISTS `totalpaytickets`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `totalpaytickets` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `NetStat` char(1) DEFAULT '',
  `Date` date NOT NULL,
  `PayMethod` int(11) DEFAULT NULL,
  `Delivered` decimal(12,2) DEFAULT '0.00',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `paytck_idx` (`Date`,`NetStat`,`PayMethod`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalpaytickets`
--

LOCK TABLES `totalpaytickets` WRITE;
/*!40000 ALTER TABLE `totalpaytickets` DISABLE KEYS */;
INSERT INTO `totalpaytickets` VALUES (1,'N','2016-10-28',0,22.69);
/*!40000 ALTER TABLE `totalpaytickets` ENABLE KEYS */;
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
) ENGINE=InnoDB AUTO_INCREMENT=513 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalscalefamilies`
--

LOCK TABLES `totalscalefamilies` WRITE;
/*!40000 ALTER TABLE `totalscalefamilies` DISABLE KEYS */;
INSERT INTO `totalscalefamilies` VALUES (1,'2016-10-01',1,1,1,1,0,0,467.838,115),(3,'2016-10-01',1,1,2,1,0,0,1194.420,122),(5,'2016-10-01',1,1,0,1,0,0,0.000,27),(6,'2016-10-01',1,1,4,1,0,0,105.856,304),(7,'2016-10-01',1,1,3,1,0,0,187.706,777),(8,'2016-10-01',1,1,466,1,0,0,0.000,1091),(11,'2016-10-01',1,1,464,1,0,0,0.000,342),(12,'2016-10-01',1,1,469,1,0,0,0.000,26),(14,'2016-10-01',1,1,465,1,0,0,2.296,362),(16,'2016-10-01',1,1,6,1,0,0,1100.840,1284),(17,'2016-10-01',1,1,9,1,0,0,175.296,607),(22,'2016-10-01',1,1,467,1,0,0,0.000,194),(31,'2016-10-01',1,1,463,1,0,0,1.755,33),(34,'2016-10-01',1,1,5,1,0,0,33.000,18),(54,'2016-10-01',1,1,462,1,0,0,0.000,61),(150,'2016-10-01',1,1,468,1,0,0,0.000,40),(168,'2016-10-01',2,2,3,1,0,0,58.783,268),(169,'2016-10-01',2,2,2,1,0,0,388.733,44),(173,'2016-10-01',2,2,4,1,0,0,25.013,90),(174,'2016-10-01',2,2,9,1,0,0,56.459,114),(176,'2016-10-01',2,2,464,1,0,0,0.000,97),(183,'2016-10-01',2,2,1,1,0,0,162.458,42),(184,'2016-10-01',2,2,6,1,0,0,380.492,401),(185,'2016-10-01',2,2,467,1,0,0,0.000,102),(187,'2016-10-01',2,2,466,1,0,0,0.000,360),(207,'2016-10-01',2,2,465,1,0,0,0.000,136),(301,'2016-10-01',2,2,463,1,0,0,0.173,26),(316,'2016-10-01',2,2,0,1,0,0,0.008,25),(332,'2016-10-01',2,2,5,1,0,0,13.633,3),(403,'2016-10-01',2,2,462,1,0,0,0.000,7),(496,'2016-10-01',2,2,469,1,0,0,0.000,11),(509,'2016-10-01',2,2,468,1,0,0,0.000,9);
/*!40000 ALTER TABLE `totalscalefamilies` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `totalvat`
--

DROP TABLE IF EXISTS `totalvat`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `totalvat` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `NetStat` char(1) DEFAULT '',
  `Date` date NOT NULL,
  `Code` decimal(1,0) unsigned NOT NULL,
  `Percentage` decimal(4,2) DEFAULT NULL,
  `Amount` decimal(12,2) DEFAULT '0.00',
  `Base` decimal(12,2) DEFAULT '0.00',
  `Tax` decimal(12,2) DEFAULT '0.00',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `dailycnt_idx` (`Date`,`NetStat`,`Code`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalvat`
--

LOCK TABLES `totalvat` WRITE;
/*!40000 ALTER TABLE `totalvat` DISABLE KEYS */;
INSERT INTO `totalvat` VALUES (1,'N','2016-10-28',2,4.00,21.19,20.37,0.82),(2,'N','2016-10-28',1,10.00,1.50,1.36,0.14);
/*!40000 ALTER TABLE `totalvat` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `trademark`
--

DROP TABLE IF EXISTS `trademark`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `trademark` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `CO_Name` varchar(255) DEFAULT '',
  `Icon` varchar(127) DEFAULT '',
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
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `trademark`
--

LOCK TABLES `trademark` WRITE;
/*!40000 ALTER TABLE `trademark` DISABLE KEYS */;
INSERT INTO `trademark` VALUES (1,'','AlejandroCespedesPantalla.png','','','','','','CASA CESPEDES','COMTE BORRELL 58','BARCELONA','08015','','93.424.28.91','','','','38104912F','','','','','','','','','','','');
/*!40000 ALTER TABLE `trademark` ENABLE KEYS */;
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
/*!50003 CREATE*/ /*!50017 DEFINER=`pcscale`@`%`*/ /*!50003 TRIGGER `TrademarkInsertedB` BEFORE INSERT ON `Trademark`
FOR EACH ROW BEGIN
  SET NEW.`Id`=1;
END */;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

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
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vats`
--

LOCK TABLES `vats` WRITE;
/*!40000 ALTER TABLE `vats` DISABLE KEYS */;
INSERT INTO `vats` VALUES (2,0,'',21.00),(3,1,'',10.00),(4,2,'',4.00),(5,3,'',0.00),(6,4,'',0.00);
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
INSERT INTO `vendorkeys` VALUES (1,1,101),(1,2,102),(1,3,103),(1,4,104),(1,5,105),(1,6,106),(1,7,107),(1,8,108),(1,9,109),(1,10,110);
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
  UPDATE `VendorRegister` SET `ReleaseDateTime`=NOW() WHERE `Vendor`=OLD.`Vendor` AND `ReleaseDateTime`=0;
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
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
  PRIMARY KEY (`Code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vendors`
--

LOCK TABLES `vendors` WRITE;
/*!40000 ALTER TABLE `vendors` DISABLE KEYS */;
INSERT INTO `vendors` VALUES (101,'Cristina','',''),(102,'Alex','',''),(103,'Eva','',''),(104,'Boris','',''),(105,'CRISTINA','',''),(106,'ALEX','',''),(107,'EVA','',''),(108,'BORIS','',''),(109,'Vendedor 109','',''),(110,'Vendedor 110','','');
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
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2016-10-28 10:24:37