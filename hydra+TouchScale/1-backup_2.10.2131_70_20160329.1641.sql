-- MySQL dump 10.14  Distrib 5.5.34-MariaDB, for debian-linux-gnu (i686)
--
-- Host: localhost    Database: pcscale
-- ------------------------------------------------------
-- Server version	5.5.34-MariaDB-1~precise-log

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
INSERT INTO `barcodeean13` VALUES (1,'G',0,'2YTTTTTIIIII','20VVViiiiiii','10VVVIIIIIII','2YTTTTTiiiii','2CCCCCCiiiii','2CCCCCCiiiii','2YTTTTTIIIII','20VVViiiiiii','10VVVIIIIIII','2YTTTTTiiiii','2CCCCCCiiiii','2CCCCCCiiiii','2CCCCCCIIIII');
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
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
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
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `counterfamilies`
--

LOCK TABLES `counterfamilies` WRITE;
/*!40000 ALTER TABLE `counterfamilies` DISABLE KEYS */;
INSERT INTO `counterfamilies` VALUES (1,1,1,101),(2,1,2,102),(3,1,3,103),(4,1,4,104),(5,1,5,105),(6,1,6,106),(7,1,7,107),(8,1,8,108),(9,1,9,109),(10,1,10,110),(11,1,11,111);
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
INSERT INTO `counters` VALUES (1,'Fruites i verdures',NULL,1,'');
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
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `customers`
--

LOCK TABLES `customers` WRITE;
/*!40000 ALTER TABLE `customers` DISABLE KEYS */;
INSERT INTO `customers` VALUES (1,'38665345Z','PEPITO','','','','','','','','','','',0.00,1);
/*!40000 ALTER TABLE `customers` ENABLE KEYS */;
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
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `erasetickets`
--

LOCK TABLES `erasetickets` WRITE;
/*!40000 ALTER TABLE `erasetickets` DISABLE KEYS */;
INSERT INTO `erasetickets` VALUES (1,'T','2016-03-29 16:17:10'),(2,'P','2016-03-29 16:17:16'),(3,'S','2016-03-29 16:17:20');
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
INSERT INTO `families` VALUES (1,'DPT. 1','','',0,''),(101,'CÍTRICS','Cítricos.jpg','',0,''),(102,'PERES','Peras.jpg','',0,''),(103,'FRUITA DE TEMPORADA','Fruta de temporada.jpg','',0,''),(104,'FRUITA TROPICAL','Frutas tropical.jpg','',0,''),(105,'TOMÀQUETS','Tomates.jpg','',0,''),(106,'POMES','Manzanas.jpg','',0,''),(107,'VERDURES','Verduras.jpg','',0,''),(108,'MELONS I SÍNDRIES','Melones y sandias.jpg','',0,''),(109,'PLÀTANS I BANANES','Platanos y bananas.jpg','',0,''),(110,'BOLETS','Setas.jpg','',0,''),(111,'SUCS','Zumos.jpeg','',0,'');
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
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `htickets`
--

LOCK TABLES `htickets` WRITE;
/*!40000 ALTER TABLE `htickets` DISABLE KEYS */;
INSERT INTO `htickets` VALUES (8,0,'N','*','0000-00-00 00:00:00','T','C',1,1,1,101,0,1,1,1.36,1.36,0.105,0,0,0.00,0,0.00,'2100001001360','2016-03-29 16:21:39',0,'2016-03-29 16:21:51',0,0,'0000-00-00 00:00:00',0,'','',0,'','0000-00-00',0.00,0,0),(9,0,'N','*','0000-00-00 00:00:00','T','C',2,1,1,101,0,1,1,1.36,1.36,0.105,0,0,0.00,0,0.00,'2100002001369','2016-03-29 16:22:14',0,'2016-03-29 16:22:21',0,0,'0000-00-00 00:00:00',0,'','',0,'','0000-00-00',0.00,0,0),(10,0,'N','*','0000-00-00 00:00:00','T','C',3,1,1,101,0,1,1,1.36,1.36,0.105,0,0,0.00,0,0.00,'2100003001368','2016-03-29 16:22:54',0,'2016-03-29 16:22:58',0,0,'0000-00-00 00:00:00',0,'','',0,'','0000-00-00',0.00,0,0),(11,0,'N','*','0000-00-00 00:00:00','T','C',4,1,2,101,0,1,1,1.30,1.30,0.100,0,0,0.00,0,0.00,'2100004001305','2016-03-29 16:23:14',0,'2016-03-29 16:23:23',0,0,'0000-00-00 00:00:00',0,'','',1,'160010100036','2016-03-29',0.00,0,0),(12,0,'N','*','0000-00-00 00:00:00','T','C',5,1,1,101,0,1,1,1.36,1.36,0.105,0,0,0.00,0,0.00,'2100005001366','2016-03-29 16:24:05',0,'2016-03-29 16:24:08',0,0,'0000-00-00 00:00:00',0,'','',0,'','0000-00-00',0.00,0,0),(13,5,'L','','0000-00-00 00:00:00','T','C',1,1,2,101,0,1,1,1.30,1.30,0.100,0,0,0.00,0,0.00,'2100001001308','2016-03-29 16:24:31',0,'2016-03-29 16:24:38',0,0,'0000-00-00 00:00:00',0,'','',0,'-No ref-','0000-00-00',0.00,0,0),(14,6,'L','','0000-00-00 00:00:00','T','C',2,1,2,101,0,1,1,0.01,0.01,0.000,1,0,0.00,0,0.00,'2100002000010','2016-03-29 16:25:10',0,'2016-03-29 16:25:18',0,0,'0000-00-00 00:00:00',0,'','',0,'','0000-00-00',0.00,0,0),(15,0,'N','*','0000-00-00 00:00:00','T','C',6,1,2,101,0,1,1,1.30,1.30,0.100,0,0,0.00,0,0.00,'2100006001303','2016-03-29 16:25:40',0,'2016-03-29 16:25:47',0,0,'0000-00-00 00:00:00',0,'','',1,'160010100037','2016-03-29',0.00,0,0);
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
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `invoicedetails`
--

LOCK TABLES `invoicedetails` WRITE;
/*!40000 ALTER TABLE `invoicedetails` DISABLE KEYS */;
INSERT INTO `invoicedetails` VALUES (3,3,11),(4,4,15);
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
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `invoices`
--

LOCK TABLES `invoices` WRITE;
/*!40000 ALTER TABLE `invoices` DISABLE KEYS */;
INSERT INTO `invoices` VALUES (3,'160010100036','','2016-03-29','','0'),(4,'160010100037','','2016-03-29','','0');
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
) ENGINE=InnoDB AUTO_INCREMENT=235 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `items`
--

LOCK TABLES `items` WRITE;
/*!40000 ALTER TABLE `items` DISABLE KEYS */;
INSERT INTO `items` VALUES (1,0,'','',NULL,NULL,NULL,NULL,NULL,NULL,0,0,999,0,0,0,0,'',0.00,0,0.00,1,0,0,0,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0),(4,5001,'XIRIMOIA','XIRIMOIA','','La xirimoia és un fruit d\'origen tropical de forma més o menys arrodonida i de pell verda reticulada, que es consumeix per la seva polpa blanca, dolça i molt aromàtica. Es pren principalment crua, encara que també s\'empra en melmelades, batuts o gelats. És un aliment ric en vitamines i minerals.\n\r\n<b>Alimentació i nutrició</b>\r\nLa xirimoia resulta útil per a persones amb problemes cardíacs, hepàtics o renals, a causa del seu alt contingut en aigua i el seu baix nivell de greixos.\nTambé pot intervenir en dietes d\'adelgazamiento, ja que conté poques calories, encara que sense oblidar que més que altres fruites carnosas. A més resulta de fàcil digestió.\n\n<b>El Fruit</b>\nLa xirimoia és una fruita tropical de forma oval, cònica, esfèrica o en forma de cor, de 7,5-12,5cm de longitud i amb un pes entre 150g i 1kg. La pell és reticulada i de color verd. La polpa blanca és sucosa, dolç, molt aromàtica i de sabor dolç alguna cosa àcid. Conté nombroses llavors aixafades de color negre de 1cm de longitud.\n\nÉs un fruit originari dels Andes, des d\'on es va estendre per Amèrica Central i alguns països Mediterranis.\n\nEs consumeix sobretot com a fruita fresca. També s\'afegeix a macedònies i sorbets i es pot usar per elaborar melmelades, batuts o gelats.\n\nLa xirimoia és un aliment ric en vitamines i minerals. Destaca pel seu contingut en vitamines A i C i en fòsfor i calci. És adequada per a persones amb problemes cardíacs, hepàtics o renals\n','','','','',104,0,1,1,0,0,0,'',1.19,0,0.00,1,1,1,1,'',0,0,0.000,'',0,'','-',0,'F_015.png','','',0,0.000,0),(5,5002,'TARONJA GRANELL RC','TARONJA GRANELL RC','','','','','','',101,0,1,1,0,0,0,'',12.99,0,0.00,1,2,1,1,'',0,0,0.000,'',0,'','-',0,'F_051.png','','',0,0.000,0),(6,5003,'LLIMONES BIO','LLIMONES BIO','','','','','','',101,0,2,1,0,0,0,'',12.99,0,0.00,1,3,1,1,'',0,0,0.000,'',0,'','-',0,'F_047.png','','',0,0.000,0),(7,5004,'LLIMONA GRANELL EC','LLIMONA GRANELL EC','','','','','','',101,0,3,1,0,0,0,'',12.99,0,0.00,1,4,1,1,'',0,0,0.000,'',0,'','-',0,'F_048.png','','',0,0.000,0),(8,5005,'MANDARINA AMB FULLA','MANDARINA AMB FULLA','','','','','','',101,0,4,1,0,0,0,'',12.99,0,0.00,1,5,1,1,'',0,0,0.000,'',0,'','-',0,'F_052.png','','',0,0.000,0),(9,5006,'MANDARINA MARCA GRANELL','MANDARINA MARCA GRANELL','','','','','','',101,0,5,1,0,0,0,'',12.99,0,0.00,1,6,1,1,'',0,0,0.000,'',0,'','-',0,'F_101.png','','',0,0.000,0),(10,5007,'ARANJA VERMELLA','ARANJA VERMELLA','','','','','','',101,0,6,1,0,0,0,'',12.99,0,0.00,1,7,1,1,'',0,0,0.000,'',0,'','-',0,'F_102.png','','',0,0.000,0),(11,5008,'TARONJA TAULA SAFATA','TARONJA TAULA SAFATA','','','','','','',101,0,7,1,0,0,0,'',12.99,0,0.00,0,8,1,1,'',0,0,0.000,'',0,'','-',0,'F_100.png','','',0,0.000,0),(12,5010,'TARONJA TAULA GRANELL','TARONJA TAULA GRANELL','','','','','','',101,0,8,1,0,0,0,'',12.99,0,0.00,1,10,1,1,'',0,0,0.000,'',0,'','-',0,'F_095.png','','',0,0.000,0),(13,5011,'TARONJA AMB FULLA','TARONJA AMB FULLA','','','','','','',101,0,9,1,0,0,0,'',12.99,0,0.00,1,11,1,1,'',0,0,0.000,'',0,'','-',0,'F_101.png','','',0,0.000,0),(14,5012,'TARONJA DIRECTA CAMP','TARONJA DIRECTA CAMP','','','','','','',101,0,10,1,0,0,0,'',12.99,0,0.00,1,12,1,1,'',0,0,0.000,'',0,'','-',0,'F_099.png','','',0,0.000,0),(15,5013,'MANDARINA SAFATA','MANDARINA SAFATA','','','','','','',101,0,11,1,0,0,0,'',12.99,0,0.00,0,13,1,1,'',0,0,0.000,'',0,'','-',0,'F_023.png','','',0,0.000,0),(16,5014,'LLIMONA SAFATA','LLIMONA SAFATA','','','','','','',101,0,12,1,0,0,0,'',12.99,0,0.00,0,14,1,1,'',0,0,0.000,'',0,'','-',0,'F_048.png','','',0,0.000,0),(17,5015,'TARONJA GRANELL (EG)','TARONJA GRANELL (EG)','','','','','','',101,0,13,1,0,0,0,'',12.99,0,0.00,1,15,1,1,'',0,0,0.000,'',0,'','-',0,'F_100.png','','',0,0.000,0),(18,5016,'ARANJA  XINESA PEÇA','ARANJA  XINESA PEÇA','','','','','','',101,0,14,1,0,0,0,'',12.99,0,0.00,0,16,1,1,'',0,0,0.000,'',0,'','-',0,'F_098.png','','',0,0.000,0),(19,5017,'MANDARINES GRANELL','MANDARINES GRANELL','','','','','','',101,0,15,1,0,0,0,'',12.99,0,0.00,1,17,1,1,'',0,0,0.000,'',0,'','-',0,'F_101.png','','',0,0.000,0),(20,5018,'MANDARINES BASIC (EG)','MANDARINES BASIC (EG)','','','','','','',101,0,16,1,0,0,0,'',12.99,0,0.00,1,18,1,1,'',0,0,0.000,'',0,'','-',0,'F_103.png','','',0,0.000,0),(21,5019,'PERA CONFERENCIA','PERA CONFERENCIA','','','','','','',102,0,1,1,0,0,0,'',12.99,0,0.00,1,19,1,1,'',0,0,0.000,'',0,'','-',0,'F_119.png','','',0,0.000,0),(22,5020,'PERA DEGANA','PERA DEGANA','','','','','','',102,0,2,1,0,0,0,'',12.99,0,0.00,1,20,1,1,'',0,0,0.000,'',0,'','-',0,'F_123.png','','',0,0.000,0),(23,5021,'POMA VERDA','POMA VERDA','','','','','','',106,0,1,1,0,0,0,'',12.99,0,0.00,1,21,1,1,'',0,0,0.000,'',0,'','-',0,'F_061.png','','',0,0.000,0),(24,5022,'PERA PASSA CRASANA','PERA PASSA CRASANA','','','','','','',102,0,3,1,0,0,0,'',12.99,0,0.00,1,22,1,1,'',0,0,0.000,'',0,'','-',0,'F_124.png','','',0,0.000,0),(25,5023,'POMA VERMELLA','POMA VERMELLA','','','','','','',106,0,2,1,0,0,0,'',12.99,0,0.00,1,23,1,1,'',0,0,0.000,'',0,'','-',0,'F_065.png','','',0,0.000,0),(26,5024,'PERA CONFERENCIA BIERZO','PERA CONFERENCIA BIERZO','','','','','','',102,0,4,1,0,0,0,'',12.99,0,0.00,1,24,1,1,'',0,0,0.000,'',0,'','-',0,'F_126.png','','',0,0.000,0),(27,5025,'POME GOLDEN FRANÇA','POME GOLDEN FRANÇA','','','','','','',106,0,3,1,0,0,0,'',12.99,0,0.00,1,25,1,1,'',0,0,0.000,'',0,'','-',0,'F_059.png','','',0,0.000,0),(28,5026,'MELÓ PARTIT','MELÓ PARTIT','','','','','','',108,0,1,1,0,0,0,'',12.99,0,0.00,1,26,1,1,'',0,0,0.000,'',0,'','-',0,'F_091.png','','',0,0.000,0),(29,5027,'POMA ROYAL GAL·LA FRANÇA','POMA ROYAL GAL·LA FRANÇA','','','','','','',106,0,4,1,0,0,0,'',12.99,0,0.00,1,27,1,1,'',0,0,0.000,'',0,'','-',0,'F_070.png','','',0,0.000,0),(30,5028,'POMA KANZI GRAN','POMA KANZI GRAN','','','','','','',106,0,5,1,0,0,0,'',12.99,0,0.00,1,28,1,1,'',0,0,0.000,'',0,'','-',0,'F_063.png','','',0,0.000,0),(31,5029,'SÍNDRIA PARTIDA','SÍNDRIA PARTIDA','','','','','','',108,0,2,1,0,0,0,'',12.99,0,0.00,1,29,1,1,'',0,0,0.000,'',0,'','-',0,'F_159.png','','',0,0.000,0),(32,5030,'BANANA GRANELL','BANANA GRANELL','','','','','','',109,0,1,1,0,0,0,'',12.99,0,0.00,1,30,1,1,'',0,0,0.000,'',0,'','-',0,'F_151.png','','',0,0.000,0),(33,5031,'PLÀTAN GRANELL','PLÀTAN GRANELL','','','','','','',109,0,2,1,0,0,0,'',12.99,0,0.00,1,31,1,1,'',0,0,0.000,'',0,'','-',0,'F_150.png','','',0,0.000,0),(34,5032,'POMA GOLDEN','POMA GOLDEN','','','','','','',106,0,6,1,0,0,0,'',12.99,0,0.00,1,32,1,1,'',0,0,0.000,'',0,'','-',0,'F_060.png','','',0,0.000,0),(35,5033,'POMA ROYAL','POMA ROYAL','','','','','','',106,0,7,1,0,0,0,'',12.99,0,0.00,1,33,1,1,'',0,0,0.000,'',0,'','-',0,'F_055.png','','',0,0.000,0),(36,5034,'PERA MAGALLON','PERA MAGALLON','','','','','','',102,0,5,1,0,0,0,'',12.99,0,0.00,1,34,1,1,'',0,0,0.000,'',0,'','-',0,'F_127.png','','',0,0.000,0),(37,5035,'POMA GRANNY (NM)','POMA GRANNY (NM)','','','','','','',106,0,8,1,0,0,0,'',12.99,0,0.00,1,35,1,1,'',0,0,0.000,'',0,'','-',0,'F_055.png','','',0,0.000,0),(38,5036,'PERA LIMONÉRA','PERA LIMONÉRA','','','','','','',102,0,6,1,0,0,0,'',12.99,0,0.00,1,36,1,1,'',0,0,0.000,'',0,'','-',0,'F_121.png','','',0,0.000,0),(39,5037,'PERA BLANQUILLA LLEIDA','PERA BLANQUILLA LLEIDA','','','','','','',102,0,7,1,0,0,0,'',12.99,0,0.00,1,37,1,1,'',0,0,0.000,'',0,'','-',0,'F_122.png','','',0,0.000,0),(40,5038,'POMA FUJI','POMA FUJI','','','','','','',106,0,9,1,0,0,0,'',12.99,0,0.00,1,38,1,1,'',0,0,0.000,'',0,'','-',0,'F_068.png','','',0,0.000,0),(41,5039,'PERA ABAT FETEL','PERA ABAT FETEL','','','','','','',102,0,8,1,0,0,0,'',12.99,0,0.00,1,39,1,1,'',0,0,0.000,'',0,'','-',0,'F_129.png','','',0,0.000,0),(42,5040,'PINYA PARTIDA','PINYA PARTIDA','','','','','','',102,0,9,1,0,0,0,'',12.99,0,0.00,1,40,1,1,'',0,0,0.000,'',0,'','-',0,'F_146.png','','',0,0.000,0),(43,5041,'POMA GOLDEN SUPREME','POMA GOLDEN SUPREME','','','','','','',106,0,10,1,0,0,0,'',12.99,0,0.00,1,41,1,1,'',0,0,0.000,'',0,'','-',0,'F_059.png','','',0,0.000,0),(44,5042,'PERA GENERAL LECLERC','PERA GENERAL LECLERC','','','','','','',102,0,10,1,0,0,0,'',12.99,0,0.00,1,42,1,1,'',0,0,0.000,'',0,'','-',0,'F_130.png','','',0,0.000,0),(45,5043,'PERA CROWN','PERA CROWN','','','','','','',102,0,11,1,0,0,0,'',12.99,0,0.00,1,43,1,1,'',0,0,0.000,'',0,'','-',0,'F_139.png','','',0,0.000,0),(46,5044,'POMA GRANNY SMITH','POMA GRANNY SMITH','','','','','','',106,0,11,1,0,0,0,'',12.99,0,0.00,1,44,1,1,'',0,0,0.000,'',0,'','-',0,'F_054.png','','',0,0.000,0),(47,5045,'POMA GOLDEN','POMA GOLDEN','','','','','','',106,0,12,1,0,0,0,'',12.99,0,0.00,1,45,1,1,'',0,0,0.000,'',0,'','-',0,'F_061.png','','',0,0.000,0),(48,5046,'POMA REINETA','POMA REINETA','','','','','','',106,0,13,1,0,0,0,'',12.99,0,0.00,1,46,1,1,'',0,0,0.000,'',0,'','-',0,'F_070.png','','',0,0.000,0),(49,5047,'POMA ROYAL GALA','POMA ROYAL GALA','','','','','','',106,0,14,1,0,0,0,'',12.99,0,0.00,1,47,1,1,'',0,0,0.000,'',0,'','-',0,'F_062.png','','',0,0.000,0),(50,5048,'POMA VERMELLA','POMA VERMELLA','','','','','','',106,0,15,1,0,0,0,'',12.99,0,0.00,1,48,1,1,'',0,0,0.000,'',0,'','-',0,'F_058.png','','',0,0.000,0),(51,5049,'PERA GENERAL LECLERC','PERA GENERAL LECLERC','','','','','','',102,0,12,1,0,0,0,'',12.99,0,0.00,1,49,1,1,'',0,0,0.000,'',0,'','-',0,'F_134.png','','',0,0.000,0),(52,5050,'BANANA GRANELL','BANANA GRANELL','','','','','','',109,0,3,1,0,0,0,'',12.99,0,0.00,1,50,1,1,'',0,0,0.000,'',0,'','-',0,'F_006.png','','',0,0.000,0),(53,5051,'SUC POMA VERMELLA','SUC POMA VERMELLA','','','','','','',111,0,1,1,0,0,0,'',12.99,0,0.00,1,51,1,1,'',0,0,0.000,'',0,'','-',0,'F_064.png','','',0,0.000,0),(54,5052,'PLÀTAN SAFATA','PLÀTAN SAFATA','','','','','','',109,0,4,1,0,0,0,'',12.99,0,0.00,0,52,1,1,'',0,0,0.000,'',0,'','-',0,'F_009.png','','',0,0.000,0),(55,5053,'POMA ROYAL SAFATA','POMA ROYAL SAFATA','','','','','','',106,0,16,1,0,0,0,'',12.99,0,0.00,0,53,1,1,'',0,0,0.000,'',0,'','-',0,'F_066.png','','',0,0.000,0),(56,5054,'POMA VERMELLA SAFATA','POMA VERMELLA SAFATA','','','','','','',106,0,17,1,0,0,0,'',12.99,0,0.00,0,54,1,1,'',0,0,0.000,'',0,'','-',0,'F_067.png','','',0,0.000,0),(57,5055,'PINYA NETA SENCERA','PINYA NETA SENCERA','','','','','','',104,0,2,1,0,0,0,'',12.99,0,0.00,1,55,1,1,'',0,0,0.000,'',0,'','-',0,'F_145.png','','',0,0.000,0),(58,5056,'POMA REINETA VERDA','POMA REINETA VERDA','','','','','','',106,0,18,1,0,0,0,'',12.99,0,0.00,1,56,1,1,'',0,0,0.000,'',0,'','-',0,'F_070.png','','',0,0.000,0),(59,5058,'PERA RED BARLETT','PERA RED BARLETT','','','','','','',102,0,13,1,0,0,0,'',12.99,0,0.00,1,58,1,1,'',0,0,0.000,'',0,'','-',0,'F_137.png','','',0,0.000,0),(60,5059,'PERA PASSA CRASSANA','PERA PASSA CRASSANA','','','','','','',102,0,14,1,0,0,0,'',12.99,0,0.00,1,59,1,1,'',0,0,0.000,'',0,'','-',0,'F_139.png','','',0,0.000,0),(61,5060,'PERA ERCOLINA GRANELL','PERA ERCOLINA GRANELL','','','','','','',102,0,15,1,0,0,0,'',12.99,0,0.00,1,60,1,1,'',0,0,0.000,'',0,'','-',0,'F_140.png','','',0,0.000,0),(62,5061,'PERA CONFERENCIA','PERA CONFERENCIA','','','','','','',102,0,16,1,0,0,0,'',12.99,0,0.00,1,61,1,1,'',0,0,0.000,'',0,'','-',0,'F_142.png','','',0,0.000,0),(63,5062,'PERA DELBARD PREMIERE','PERA DELBARD PREMIERE','','','','','','',102,0,17,1,0,0,0,'',12.99,0,0.00,1,62,1,1,'',0,0,0.000,'',0,'','-',0,'F_132.png','','',0,0.000,0),(64,5063,'PERA CONFERENCIA PREMIUM','PERA CONFERENCIA PREMIUM','','','','','','',102,0,18,1,0,0,0,'',12.99,0,0.00,0,63,1,1,'',0,0,0.000,'',0,'','-',0,'F_126.png','','',0,0.000,0),(65,5064,'SUC TARONJA NATURAL 250cc','SUC TARONJA NATURAL 250cc','','','','','','',111,0,2,1,0,0,0,'',12.99,0,0.00,0,64,1,1,'',0,0,0.000,'',0,'','-',0,'F_177.png','','',0,0.000,0),(66,5065,'POMA REINETA GRIS','POMA REINETA GRIS','','','','','','',106,0,19,1,0,0,0,'',12.99,0,0.00,1,65,1,1,'',0,0,0.000,'',0,'','-',0,'F_061.png','','',0,0.000,0),(67,5066,'POMA ROYAL BIO','POMA ROYAL BIO','','','','','','',106,0,20,1,0,0,0,'',12.99,0,0.00,1,66,1,1,'',0,0,0.000,'',0,'','-',0,'F_054.png','','',0,0.000,0),(68,5067,'POMA GOLDEN BIO','POMA GOLDEN BIO','','','','','','',106,0,21,1,0,0,0,'',12.99,0,0.00,1,67,1,1,'',0,0,0.000,'',0,'','-',0,'F_060.png','','',0,0.000,0),(69,5069,'POMA ROYAL GALA','POMA ROYAL GALA','','','','','','',106,0,22,1,0,0,0,'',12.99,0,0.00,1,69,1,1,'',0,0,0.000,'',0,'','-',0,'F_061.png','','',0,0.000,0),(70,5070,'POMA GOLDEN SAFATA','POMA GOLDEN SAFATA','','','','','','',106,0,23,1,0,0,0,'',12.99,0,0.00,0,70,1,1,'',0,0,0.000,'',0,'','-',0,'F_059.png','','',0,0.000,0),(71,5071,'CONFERENCIA (EG)','CONFERENCIA (EG)','','','','','','',102,0,19,1,0,0,0,'',12.99,0,0.00,1,71,1,1,'',0,0,0.000,'',0,'','-',0,'F_125.png','','',0,0.000,0),(72,5072,'POMA FUJI (NM)','POMA FUJI (NM)','','','','','','',106,0,24,1,0,0,0,'',12.99,0,0.00,1,72,1,1,'',0,0,0.000,'',0,'','-',0,'F_055.png','','',0,0.000,0),(73,5073,'PLÀTAN 1 BORSA','PLÀTAN 1 BORSA','','','','','','',109,0,5,1,0,0,0,'',12.99,0,0.00,1,73,1,1,'',0,0,0.000,'',0,'','-',0,'F_009.png','','',0,0.000,0),(74,5074,'BLANQUILLA (EG)','BLANQUILLA (EG)','','','','','','',102,0,20,1,0,0,0,'',12.99,0,0.00,1,74,1,1,'',0,0,0.000,'',0,'','-',0,'F_120.png','','',0,0.000,0),(75,5075,'PERA NASHI','PERA NASHI','','','','','','',102,0,21,1,0,0,0,'',12.99,0,0.00,1,75,1,1,'',0,0,0.000,'',0,'','-',0,'F_131.png','','',0,0.000,0),(76,5076,'PERA ALEXANDRINA','PERA ALEXANDRINA','','','','','','',102,0,22,1,0,0,0,'',12.99,0,0.00,1,76,1,1,'',0,0,0.000,'',0,'','-',0,'F_128.png','','',0,0.000,0),(77,5077,'POMA GOLDEN','POMA GOLDEN','','','','','','',106,0,25,1,0,0,0,'',12.99,0,0.00,1,77,1,1,'',0,0,0.000,'',0,'','-',0,'F_060.png','','',0,0.000,0),(78,5078,'PERA PACKAMS','PERA PACKAMS','','','','','','',102,0,23,1,0,0,0,'',12.99,0,0.00,1,78,1,1,'',0,0,0.000,'',0,'','-',0,'F_138.png','','',0,0.000,0),(79,5079,'POMA VERMELLA GRANELL','POMA VERMELLA GRANELL','','','','','','',106,0,26,1,0,0,0,'',12.99,0,0.00,1,79,1,1,'',0,0,0.000,'',0,'','-',0,'F_067.png','','',0,0.000,0),(80,5080,'POMA CANADA','POMA CANADA','','','','','','',106,0,27,1,0,0,0,'',12.99,0,0.00,1,80,1,1,'',0,0,0.000,'',0,'','-',0,'F_063.png','','',0,0.000,0),(81,5081,'PERA ROTXA','PERA ROTXA','','','','','','',102,0,24,1,0,0,0,'',12.99,0,0.00,1,81,1,1,'',0,0,0.000,'',0,'','-',0,'F_140.png','','',0,0.000,0),(82,5082,'COL ARRISSAT','COL ARRISSAT','','','','','','',107,0,1,1,0,0,0,'',12.99,0,0.00,1,82,1,1,'',0,0,0.000,'',0,'','-',0,'V_038.png','','',0,0.000,0),(83,5083,'PERA CONFERENCIA PREMIUM','PERA CONFERENCIA PREMIUM SAFATA','','','','','','',102,0,25,1,0,0,0,'',12.99,0,0.00,1,83,1,1,'',0,0,0.000,'',0,'','-',0,'F_119.png','','',0,0.000,0),(84,5084,'PERA BLANQUILLA','PERA BLANQUILLA','','','','','','',102,0,26,1,0,0,0,'',12.99,0,0.00,1,84,1,1,'',0,0,0.000,'',0,'','-',0,'F_122.png','','',0,0.000,0),(85,5085,'MACEDONIA FRUTAS DEL TIEM','MACEDONIA FRUTAS DEL TIEMPO','','','','','','',103,0,1,1,0,0,0,'',12.99,0,0.00,0,85,1,1,'',0,0,0.000,'',0,'','-',0,'F_175.png','','',0,0.000,0),(86,5086,'PERA CONFERENCIA BIERZO (','PERA CONFERENCIA BIERZO (EG)','','','','','','',102,0,27,1,0,0,0,'',12.99,0,0.00,1,86,1,1,'',0,0,0.000,'',0,'','-',0,'F_123.png','','',0,0.000,0),(87,5087,'PLÀTAN EXTRA CANÀRIES','PLÀTAN EXTRA CANÀRIES','','','','','','',109,0,6,1,0,0,0,'',12.99,0,0.00,1,87,1,1,'',0,0,0.000,'',0,'','-',0,'F_008.png','','',0,0.000,0),(88,5091,'MELÓ VERD PARTIT','MELÓ VERD PARTIT','','','','','','',108,0,3,1,0,0,0,'',12.99,0,0.00,1,91,1,1,'',0,0,0.000,'',0,'','-',0,'F_087.png','','',0,0.000,0),(89,5096,'BANANA 1 BORSA','BANANA 1 BORSA','','','','','','',109,0,7,1,0,0,0,'',12.99,0,0.00,1,96,1,1,'',0,0,0.000,'',0,'','-',0,'F_006.png','','',0,0.000,0),(90,5098,'SUC ARANJA 500cc','SUC ARANJA 500cc','','','','','','',111,0,3,1,0,0,0,'',12.99,0,0.00,0,98,1,1,'',0,0,0.000,'',0,'','-',0,'F_177.png','','',0,0.000,0),(91,5099,'NISPERO GRANEL (EG)','NISPERO GRANEL (EG)','','','','','','',103,0,2,1,0,0,0,'',12.99,0,0.00,1,99,1,1,'',0,0,0.000,'',0,'','-',0,'F_107.png','','',0,0.000,0),(92,5100,'KAKI SHARONI','KAKI SHARONI','','','','','','',103,0,3,1,0,0,0,'',12.99,0,0.00,1,100,1,1,'',0,0,0.000,'',0,'','-',0,'F_011.png','','',0,0.000,0),(93,5101,'SÍNDRIA NEGRA','SÍNDRIA NEGRA','','','','','','',108,0,4,1,0,0,0,'',12.99,0,0.00,1,101,1,1,'',1,0,0.000,'',0,'','-',0,'F_158.png','','',0,0.000,0),(94,5102,'MELÓ CANTALOUP','MELÓ CANTALOUP','','','','','','',108,0,5,1,0,0,0,'',12.99,0,0.00,1,102,1,1,'',1,0,0.000,'',0,'','-',0,'F_083.png','','',0,0.000,0),(95,5103,'ALBERCOC ARAGÓ','ALBERCOC ARAGÓ','','','','','','',103,0,4,1,0,0,0,'',12.99,0,0.00,1,103,1,1,'',0,0,0.000,'',0,'','-',0,'F_080.png','','',0,0.000,0),(96,5104,'MELÓ GROC','MELÓ GROC','','','','','','',108,0,6,1,0,0,0,'',12.99,0,0.00,1,104,1,1,'',3,0,0.000,'',0,'','-',0,'F_084.png','','',0,0.000,0),(97,5105,'SÍNDRIA RATJADA','SÍNDRIA RATJADA','','','','','','',108,0,7,1,0,0,0,'',12.99,0,0.00,1,105,1,1,'',1,0,0.000,'',0,'','-',0,'F_160.png','','',0,0.000,0),(98,5106,'MELÓ CHARENTAIS','MELÓ CHARENTAIS','','','','','','',108,0,8,1,0,0,0,'',12.99,0,0.00,1,106,1,1,'',0,0,0.000,'',0,'','-',0,'F_165.png','','',0,0.000,0),(99,5107,'ALBERCOC','ALBERCOC','','','','','','',103,0,5,1,0,0,0,'',12.99,0,0.00,1,107,1,1,'',0,0,0.000,'',0,'','-',0,'F_080.png','','',0,0.000,0),(100,5108,'PRÉSSEC GROC SELECCIÓ','PRÉSSEC GROC SELECCIÓ','','','','','','',103,0,6,1,0,0,0,'',12.99,0,0.00,1,108,1,1,'',0,0,0.000,'',0,'','-',0,'F_076.png','','',0,0.000,0),(101,5109,'PRÉSSEC GROC','PRÉSSEC GROC','','','','','','',103,0,7,1,0,0,0,'',12.99,0,0.00,1,109,1,1,'',0,0,0.000,'',0,'','-',0,'F_075.png','','',0,0.000,0),(102,5110,'PRUNA VERMELLA (EG)','PRUNA VERMELLA (EG)','','','','','','',103,0,8,1,0,0,0,'',12.99,0,0.00,1,110,1,1,'',0,0,0.000,'',0,'','-',0,'F_018.png','','',0,0.000,0),(103,5111,'PRÈSSEC VERMELL','PRÈSSEC VERMELL','','','','','','',103,0,9,1,0,0,0,'',12.99,0,0.00,1,111,1,1,'',0,0,0.000,'',0,'','-',0,'F_031.png','','',0,0.000,0),(104,5112,'CIRERA IMPERIAL GRANELL','CIRERA IMPERIAL GRANELL','','','','','','',103,0,10,1,0,0,0,'',12.99,0,0.00,1,112,1,1,'',0,0,0.000,'',0,'','-',0,'F_143.png','','',0,0.000,0),(105,5113,'NECTARINA ARAGÓ','NECTARINA ARAGÓ','','','','','','',103,0,11,1,0,0,0,'',12.99,0,0.00,1,113,1,1,'',0,0,0.000,'',0,'','-',0,'F_098.png','','',0,0.000,0),(106,5114,'PARAGUAIÀ LLEIDA','PARAGUAIÀ LLEIDA','','','','','','',103,0,12,1,0,0,0,'',12.99,0,0.00,1,114,1,1,'',0,0,0.000,'',0,'','-',0,'F_113.png','','',0,0.000,0),(107,5115,'FIGUES NEGRES','FIGUES NEGRES','','','','','','',103,0,13,1,0,0,0,'',12.99,0,0.00,1,115,1,1,'',0,0,0.000,'',0,'','-',0,'F_036.png','','',0,0.000,0),(108,5116,'MELÓ PELL GRIPAU TALL','MELÓ PELL GRIPAU TALL','','','','','','',103,0,14,1,0,0,0,'',12.99,0,0.00,0,116,1,1,'',1,0,0.000,'',0,'','-',0,'F_086.png','','',0,0.000,0),(109,5117,'PRÉSSEC','PRÉSSEC','','','','','','',103,0,15,1,0,0,0,'',12.99,0,0.00,1,117,1,1,'',0,0,0.000,'',0,'','-',0,'F_080.png','','',0,0.000,0),(110,5118,'MELÓ GÀL·LIA','MELÓ GÀL·LIA','','','','','','',108,0,9,1,0,0,0,'',12.99,0,0.00,1,118,1,1,'',1,0,0.000,'',0,'','-',0,'F_089.png','','',0,0.000,0),(111,5119,'PARAGUAIÀ','PARAGUAIÀ','','','','','','',103,0,16,1,0,0,0,'',12.99,0,0.00,1,119,1,1,'',0,0,0.000,'',0,'','-',0,'F_113.png','','',0,0.000,0),(112,5120,'PRÉSSEC CALANDA','PRÉSSEC CALANDA','','','','','','',103,0,17,1,0,0,0,'',12.99,0,0.00,1,120,1,1,'',0,0,0.000,'',0,'','-',0,'F_079.png','','',0,0.000,0),(113,5121,'PRÉSSEC GROC','PRÉSSEC GROC','','','','','','',103,0,18,1,0,0,0,'',12.99,0,0.00,1,121,1,1,'',0,0,0.000,'',0,'','-',0,'F_080.png','','',0,0.000,0),(114,5122,'MELÓ PELL GRIPAU','MELÓ PELL GRIPAU','','','','','','',103,0,19,1,0,0,0,'',12.99,0,0.00,1,122,1,1,'',0,0,0.000,'',0,'','-',0,'F_084.png','','',0,0.000,0),(115,5123,'RAÏM MUSCAT GRANELL','RAÏM MUSCAT GRANELL','','','','','','',103,0,20,1,0,0,0,'',12.99,0,0.00,1,123,1,1,'',0,0,0.000,'',0,'','-',0,'F_167.png','','',0,0.000,0),(116,5124,'RAÏM CHASSELAS','RAÏM CHASSELAS','','','','','','',103,0,21,1,0,0,0,'',12.99,0,0.00,1,124,1,1,'',0,0,0.000,'',0,'','-',0,'F_168.png','','',0,0.000,0),(117,5125,'PÉSSEC VERMMEL (NM)','PÉSSEC VERMMEL (NM)','','','','','','',103,0,22,1,0,0,0,'',12.99,0,0.00,1,125,1,1,'',0,0,0.000,'',0,'','-',0,'F_079.png','','',0,0.000,0),(118,5126,'RAÏM ROSAT','RAÏM ROSAT','','','','','','',103,0,23,1,0,0,0,'',12.99,0,0.00,1,126,1,1,'',0,0,0.000,'',0,'','-',0,'F_173.png','','',0,0.000,0),(119,5127,'PRÉSSEC VERMELL (EG)','PRÉSSEC VERMELL (EG)','','','','','','',103,0,24,1,0,0,0,'',12.99,0,0.00,1,127,1,1,'',0,0,0.000,'',0,'','-',0,'F_031.png','','',0,0.000,0),(120,5128,'MAGRANA','MAGRANA','','','','','','',103,0,25,1,0,0,0,'',12.99,0,0.00,1,128,1,1,'',0,0,0.000,'',0,'','-',0,'F_033.png','','',0,0.000,0),(121,5129,'PLATERINA GRANELL','PLATERINA GRANELL','','','','','','',103,0,26,1,0,0,0,'',12.99,0,0.00,1,129,1,1,'',0,0,0.000,'',0,'','-',0,'F_105.png','','',0,0.000,0),(122,5130,'PRUNA GROGA GRANELL','PRUNA GROGA GRANELL','','','','','','',103,0,27,1,0,0,0,'',12.99,0,0.00,1,130,1,1,'',0,0,0.000,'',0,'','-',0,'F_019.png','','',0,0.000,0),(123,5131,'NECTARINA GROGA','NECTARINA GROGA','','','','','','',103,0,28,1,0,0,0,'',12.99,0,0.00,1,131,1,1,'',0,0,0.000,'',0,'','-',0,'F_053.png','','',0,0.000,0),(124,5132,'FIGA DE MORO','FIGA DE MORO','','','','','','',103,0,29,1,0,0,0,'',12.99,0,0.00,1,132,1,1,'',0,0,0.000,'',0,'','-',0,'F_035.png','','',0,0.000,0),(125,5133,'CASTANYA GRANELL','CASTANYA GRANELL','','','','','','',103,0,30,1,0,0,0,'',12.99,0,0.00,1,133,1,1,'',0,0,0.000,'',0,'','-',0,'FS_020.png','','',0,0.000,0),(126,5134,'NECTARINA CARN BLANCA','NECTARINA CARN BLANCA','','','','','','',103,0,31,1,0,0,0,'',12.99,0,0.00,1,134,1,1,'',0,0,0.000,'',0,'','-',0,'F_053.png','','',0,0.000,0),(127,5135,'CODONY','CODONY','','','','','','',103,0,32,1,0,0,0,'',12.99,0,0.00,1,135,1,1,'',0,0,0.000,'',0,'','-',0,'F_166.png','','',0,0.000,0),(128,5136,'KAKI','KAKI','','','','','','',103,0,33,1,0,0,0,'',12.99,0,0.00,1,136,1,1,'',0,0,0.000,'',0,'','-',0,'F_011.png','','',0,0.000,0),(129,5137,'PRUNA REGNA CLAUDIA VERD','PRUNA REGNA CLAUDIA VERD','','','','','','',103,0,34,1,0,0,0,'',12.99,0,0.00,1,137,1,1,'',0,0,0.000,'',0,'','-',0,'F_020.png','','',0,0.000,0),(130,5138,'NECTARINA GROGA (EG)','NECTARINA GROGA (EG)','','','','','','',103,0,35,1,0,0,0,'',12.99,0,0.00,1,138,1,1,'',0,0,0.000,'',0,'','-',0,'F_053.png','','',0,0.000,0),(131,5139,'CASTANYES DE LLEÓ','CASTANYES DE LLEÓ','','','','','','',103,0,36,1,0,0,0,'',12.99,0,0.00,1,139,1,1,'',0,0,0.000,'',0,'','-',0,'FS_019.png','','',0,0.000,0),(132,5140,'RAÏM BLANC','RAÏM BLANC','','','','','','',103,0,37,1,0,0,0,'',12.99,0,0.00,1,140,1,1,'',0,0,0.000,'',0,'','-',0,'F_174.png','','',0,0.000,0),(133,5141,'MELÓ PELL DE GRIPAU MADRI','MELÓ PELL DE GRIPAU MADRID','','','','','','',108,0,10,1,0,0,0,'',12.99,0,0.00,1,141,1,1,'',0,0,0.000,'',0,'','-',0,'F_085.png','','',0,0.000,0),(134,5142,'MADUIXOT GRANELL (EG)','MADUIXOT GRANELL (EG)','','','','','','',103,0,38,1,0,0,0,'',12.99,0,0.00,1,142,1,1,'',0,0,0.000,'',0,'','-',0,'F_030.png','','',0,0.000,0),(135,5143,'PRÉSSEC GROC SAFATA','PRÉSSEC GROC SAFATA','','','','','','',103,0,39,1,0,0,0,'',12.99,0,0.00,0,143,1,1,'',0,0,0.000,'',0,'','-',0,'F_002.png','','',0,0.000,0),(136,5144,'RAÏM NEGRE','RAÏM NEGRE','','','','','','',103,0,40,1,0,0,0,'',12.99,0,0.00,1,144,1,1,'',0,0,0.000,'',0,'','-',0,'F_171.png','','',0,0.000,0),(137,5145,'PRÉSSEC GROC (EG)','PRÉSSEC GROC (EG)','','','','','','',103,0,41,1,0,0,0,'',12.99,0,0.00,1,145,1,1,'',0,0,0.000,'',0,'','-',0,'F_002.png','','',0,0.000,0),(138,5146,'CIRERA GRANELL RC','CIRERA GRANELL RC','','','','','','',103,0,42,1,0,0,0,'',12.99,0,0.00,1,146,1,1,'',0,0,0.000,'',0,'','-',0,'F_143.png','','',0,0.000,0),(139,5147,'RAÏM BLANC SAFATA','RAÏM BLANC SAFATA','','','','','','',103,0,43,1,0,0,0,'',12.99,0,0.00,0,147,1,1,'',0,0,0.000,'',0,'','-',0,'F_167.png','','',0,0.000,0),(140,5148,'SÍNDRIA','SÍNDRIA','','','','','','',108,0,11,1,0,0,0,'',12.99,0,0.00,1,148,1,1,'',1,0,0.000,'',0,'','-',0,'F_162.png','','',0,0.000,0),(141,5151,'PLÀTAN DE FREGIR','PLÀTAN DE FREGIR','','','','','','',109,0,8,1,0,0,0,'',12.99,0,0.00,1,151,1,1,'',0,0,0.000,'',0,'','-',0,'F_152.png','','',0,0.000,0),(142,5152,'YUCA','YUCA','','','','','','',104,0,3,1,0,0,0,'',12.99,0,0.00,1,152,1,1,'',0,0,0.000,'',0,'','-',0,'F_178.png','','',0,0.000,0),(143,5153,'XIRIMOIA SAFATA','XIRIMOIA SAFATA','','','','','','',104,0,4,1,0,0,0,'',12.99,0,0.00,0,153,1,1,'',0,0,0.000,'',0,'','-',0,'F_016.png','','',0,0.000,0),(144,5154,'LLIMA','LLIMA','','','','','','',104,0,5,1,0,0,0,'',12.99,0,0.00,1,154,1,1,'',0,0,0.000,'',0,'','-',0,'F_046.png','','',0,0.000,0),(145,5155,'PEBROT VERD','PEBROT VERD','','','','','','',107,0,2,1,0,0,0,'',12.99,0,0.00,1,155,1,1,'',0,0,0.000,'',0,'','-',0,'V_082.png','','',0,0.000,0),(146,5156,'LITXIS','LITXIS','','','','','','',104,0,6,1,0,0,0,'',12.99,0,0.00,1,156,1,1,'',0,0,0.000,'',0,'','-',0,'F_049.png','','',0,0.000,0),(147,5157,'PINYA','PINYA','','','','','','',104,0,7,1,0,0,0,'',12.99,0,0.00,1,157,1,1,'',1,0,0.000,'',0,'','-',0,'F_147.png','','',0,0.000,0),(148,5158,'KIWI ZESPRI GRANELL','KIWI ZESPRI GRANELL','','','','','','',104,0,8,1,0,0,0,'',12.99,0,0.00,1,158,1,1,'',0,0,0.000,'',0,'','-',0,'F_043.png','','',0,0.000,0),(149,5159,'FRUITA PASIÓ','FRUITA PASIÓ','','','','','','',104,0,9,1,0,0,0,'',12.99,0,0.00,1,159,1,1,'',0,0,0.000,'',0,'','-',0,'F_074.png','','',0,0.000,0),(150,5160,'LLIMA','LLIMA','','','','','','',104,0,10,1,0,0,0,'',12.99,0,0.00,1,160,1,1,'',0,0,0.000,'',0,'','-',0,'F_046.png','','',0,0.000,0),(151,5161,'TAMARILLO','TAMARILLO','','','','','','',104,0,11,1,0,0,0,'',12.99,0,0.00,1,161,1,1,'',0,0,0.000,'',0,'','-',0,'F_033.png','','',0,0.000,0),(152,5162,'MÀNEC','MÀNEC','','','','','','',104,0,12,1,0,0,0,'',12.99,0,0.00,1,162,1,1,'',0,0,0.000,'',0,'','-',0,'F_110.png','','',0,0.000,0),(153,5163,'PANSES','PANSES','','','','','','',104,0,13,1,0,0,0,'',12.99,0,0.00,1,163,1,1,'',0,0,0.000,'',0,'','-',0,'F_022.png','','',0,0.000,0),(154,5164,'GINGEBRE','GINGEBRE','','','','','','',104,0,14,1,0,0,0,'',12.99,0,0.00,1,164,1,1,'',0,0,0.000,'',0,'','-',0,'V_056.png','','',0,0.000,0),(155,5165,'PINYA GOLDEN RIPE','PINYA GOLDEN RIPE','','','','','','',104,0,15,1,0,0,0,'',12.99,0,0.00,1,165,1,1,'',1,0,0.000,'',0,'','-',0,'F_148.png','','',0,0.000,0),(156,5166,'GUAIABA','GUAIABA','','','','','','',104,0,16,1,0,0,0,'',12.99,0,0.00,1,166,1,1,'',0,0,0.000,'',0,'','-',0,'F_111.png','','',0,0.000,0),(157,5167,'PINYA BABY','PINYA BABY','','','','','','',104,0,17,1,0,0,0,'',12.99,0,0.00,0,167,1,1,'',0,0,0.000,'',0,'','-',0,'F_146.png','','',0,0.000,0),(158,5168,'KIWI GOLD GRANELL','KIWI GOLD GRANELL','','','','','','',104,0,18,1,0,0,0,'',12.99,0,0.00,1,168,1,1,'',0,0,0.000,'',0,'','-',0,'F_044.png','','',0,0.000,0),(159,5169,'ALVOCAT HASS','ALVOCAT HASS','','','','','','',104,0,19,1,0,0,0,'',12.99,0,0.00,1,169,1,1,'',0,0,0.000,'',0,'','-',0,'F_034.png','','',0,0.000,0),(160,5170,'KIWI (EG)','KIWI (EG)','','','','','','',104,0,20,1,0,0,0,'',12.99,0,0.00,1,170,1,1,'',0,0,0.000,'',0,'','-',0,'F_044.png','','',0,0.000,0),(161,5171,'PHYSALIS 100g','PHYSALIS 100g','','','','','','',104,0,21,1,0,0,0,'',12.99,0,0.00,0,171,1,1,'',0,0,0.000,'',0,'','-',0,'F_156.png','','',0,0.000,0),(162,5172,'ALVOCAT LLIS','ALVOCAT LLIS','','','','','','',104,0,22,1,0,0,0,'',12.99,0,0.00,1,172,1,1,'',0,0,0.000,'',0,'','-',0,'F_034.png','','',0,0.000,0),(163,5173,'PAPAIA','PAPAIA','','','','','','',104,0,23,1,0,0,0,'',12.99,0,0.00,1,173,1,1,'',0,0,0.000,'',0,'','-',0,'F_110.png','','',0,0.000,0),(164,5174,'PASSIONERA','PASSIONERA','','','','','','',104,0,24,1,0,0,0,'',12.99,0,0.00,1,174,1,1,'',0,0,0.000,'',0,'','-',0,'F_116.png','','',0,0.000,0),(165,5175,'COCO','COCO','','','','','','',104,0,25,1,0,0,0,'',12.99,0,0.00,0,175,1,1,'',0,0,0.000,'',0,'','-',0,'F_024.png','','',0,0.000,0),(166,5176,'ALVOCAT GRANELL','ALVOCAT GRANELL','','','','','','',104,0,26,1,0,0,0,'',12.99,0,0.00,1,176,1,1,'',0,0,0.000,'',0,'','-',0,'F_034.png','','',0,0.000,0),(167,5177,'TOMÀQUET GUST AMANIDA','TOMÀQUET GUST AMANIDA','','','','','','',105,0,1,1,0,0,0,'',12.99,0,0.00,1,177,1,1,'',0,0,0.000,'',0,'','-',0,'V_103.png','','',0,0.000,0),(168,5178,'TOMÀQUET MADUR BOLA','TOMÀQUET MADUR BOLA','','','','','','',105,0,2,1,0,0,0,'',12.99,0,0.00,1,178,1,1,'',0,0,0.000,'',0,'','-',0,'V_102.png','','',0,0.000,0),(169,5179,'TOMÀQUET AMANIDA','TOMÀQUET AMANIDA','','','','','','',105,0,3,1,0,0,0,'',12.99,0,0.00,1,179,1,1,'',0,0,0.000,'',0,'','-',0,'V_109.png','','',0,0.000,0),(170,5180,'TOMÀQUET AMANIDA SAFATA','TOMÀQUET AMANIDA SAFATA','','','','','','',105,0,4,1,0,0,0,'',12.99,0,0.00,0,180,1,1,'',0,0,0.000,'',0,'','-',0,'V_105.png','','',0,0.000,0),(171,5181,'TOMÀQUET COR BOU','TOMÀQUET COR BOU','','','','','','',105,0,5,1,0,0,0,'',12.99,0,0.00,1,181,1,1,'',0,0,0.000,'',0,'','-',0,'V_106.png','','',0,0.000,0),(172,5182,'TOMÀQUET RAFFELLO GRANELL','TOMÀQUET RAFFELLO GRANELL','','','','','','',105,0,6,1,0,0,0,'',12.99,0,0.00,1,182,1,1,'',0,0,0.000,'',0,'','-',0,'V_105.png','','',0,0.000,0),(173,5183,'TOMÀQUET CHERRY MINI KUMA','TOMÀQUET CHERRY MINI KUMATO','','','','','','',105,0,7,1,0,0,0,'',12.99,0,0.00,1,183,1,1,'',0,0,0.000,'',0,'','-',0,'V_110.png','','',0,0.000,0),(174,5184,'CEP (BOLETS)','CEP (BOLETS)','','','','','','',110,0,1,1,0,0,0,'',12.99,0,0.00,1,184,1,1,'',0,0,0.000,'',0,'','-',0,'V_099.png','','',0,0.000,0),(175,5185,'TOMÀQUET BOLA SAFATA','TOMÀQUET BOLA SAFATA','','','','','','',105,0,8,1,0,0,0,'',12.99,0,0.00,0,185,1,1,'',0,0,0.000,'',0,'','-',0,'V_113.png','','',0,0.000,0),(176,5186,'TOMÀQUET KUMATO','TOMÀQUET KUMATO','','','','','','',105,0,9,1,0,0,0,'',12.99,0,0.00,1,186,1,1,'',0,0,0.000,'',0,'','-',0,'V_107.png','','',0,0.000,0),(177,5187,'PASTANAGA GRANELL','PASTANAGA GRANELL','','','','','','',107,0,3,1,0,0,0,'',12.99,0,0.00,1,187,1,1,'',0,0,0.000,'',0,'','-',0,'V_115.png','','',0,0.000,0),(178,5188,'TOMÀQUET GUST PEBROT','TOMÀQUET GUST PEBROT','','','','','','',105,0,10,1,0,0,0,'',12.99,0,0.00,1,188,1,1,'',0,0,0.000,'',0,'','-',0,'V_112.png','','',0,0.000,0),(179,5189,'MONGETA AMPLA GRANELL','MONGETA AMPLA GRANELL','','','','','','',107,0,4,1,0,0,0,'',12.99,0,0.00,1,189,1,1,'',0,0,0.000,'',0,'','-',0,'V_101.png','','',0,0.000,0),(180,5190,'CAMA-SEC (BOLET)','CAMA-SEC (BOLET)','','','','','','',110,0,2,1,0,0,0,'',12.99,0,0.00,1,190,1,1,'',0,0,0.000,'',0,'','-',0,'V_098.png','','',0,0.000,0),(181,5191,'MONGETA PAIMPOL','MONGETA PAIMPOL','','','','','','',107,0,5,1,0,0,0,'',12.99,0,0.00,1,191,1,1,'',0,0,0.000,'',0,'','-',0,'V_053.png','','',0,0.000,0),(182,5192,'TOMÀQUET CHERRY GRAN','TOMÀQUET CHERRY GRAN','','','','','','',105,0,11,1,0,0,0,'',12.99,0,0.00,1,192,1,1,'',0,0,0.000,'',0,'','-',0,'V_104.png','','',0,0.000,0),(183,5193,'CARBASSA SEPAIO','CARBASSA SEPAIO','','','','','','',107,0,6,1,0,0,0,'',12.99,0,0.00,1,193,1,1,'',4,0,0.000,'',0,'','-',0,'V_026.png','','',0,0.000,0),(184,5194,'ROSSINYOL GRANELL','ROSSINYOL GRANELL','','','','','','',110,0,3,1,0,0,0,'',12.99,0,0.00,1,194,1,1,'',0,0,0.000,'',0,'','-',0,'V_098.png','','',0,0.000,0),(185,5195,'BOLET ORELLA JUDES','BOLET ORELLA JUDES','','','','','','',110,0,4,1,0,0,0,'',12.99,0,0.00,1,195,1,1,'',0,0,0.000,'',0,'','-',0,'V_099.png','','',0,0.000,0),(186,5196,'BOLET ANGULA DE MUNTANYA','BOLET ANGULA DE MUNTANYA','','','','','','',110,0,5,1,0,0,0,'',12.99,0,0.00,1,196,1,1,'',0,0,0.000,'',0,'','-',0,'V_073.png','','',0,0.000,0),(187,5197,'TOMÀQUET ENSALADA MADRILE','TOMÀQUET ENSALADA MADRILENYA','','','','','','',105,0,12,1,0,0,0,'',12.99,0,0.00,1,197,1,1,'',0,0,0.000,'',0,'','-',0,'V_109.png','','',0,0.000,0),(188,5198,'TOMÀQUET CHERRY PERA','TOMÀQUET CHERRY PERA','','','','','','',105,0,13,1,0,0,0,'',12.99,0,0.00,1,198,1,1,'',0,0,0.000,'',0,'','-',0,'V_111.png','','',0,0.000,0),(189,5199,'MONGETA FINA GRANELL','MONGETA FINA GRANELL','','','','','','',107,0,7,1,0,0,0,'',12.99,0,0.00,1,199,1,1,'',0,0,0.000,'',0,'','-',0,'V_059.png','','',0,0.000,0),(190,5200,'TOMÀQUET CHERRY AMARG','TOMÀQUET CHERRY AMARG','','','','','','',105,0,14,1,0,0,0,'',12.99,0,0.00,1,200,1,1,'',0,0,0.000,'',0,'','-',0,'F_156.png','','',0,0.000,0),(191,5201,'CARABASSA PER ROSTIR','CARABASSA PER ROSTIR','','','','','','',107,0,8,1,0,0,0,'',12.99,0,0.00,1,201,1,1,'',0,0,0.000,'',0,'','-',0,'V_026.png','','',0,0.000,0),(192,5202,'TOMÀQUET CHERRY GRAN','TOMÀQUET CHERRY GRAN','','','','','','',105,0,15,1,0,0,0,'',12.99,0,0.00,1,202,1,1,'',0,0,0.000,'',0,'','-',0,'V_104.png','','',0,0.000,0),(193,5203,'COGOMBRE FRANCÈS CAMP','COGOMBRE FRANCÈS CAMP','','','','','','',107,0,9,1,0,0,0,'',12.99,0,0.00,1,203,1,1,'',0,0,0.000,'',0,'','-',0,'V_075.png','','',0,0.000,0),(194,5204,'TOMÀQUET BRANCA','TOMÀQUET BRANCA','','','','','','',105,0,16,1,0,0,0,'',12.99,0,0.00,1,204,1,1,'',0,0,0.000,'',0,'','-',0,'V_108.png','','',0,0.000,0),(195,5205,'TOMÀQUET BRANCA CAMP','TOMÀQUET BRANCA CAMP','','','','','','',105,0,17,1,0,0,0,'',12.99,0,0.00,1,205,1,1,'',0,0,0.000,'',0,'','-',0,'V_103.png','','',0,0.000,0),(196,5206,'ALBERGÍNIA SORAIA','ALBERGÍNIA SORAIA','','','','','','',107,0,10,1,0,0,0,'',12.99,0,0.00,1,206,1,1,'',0,0,0.000,'',0,'','-',0,'V_012.png','','',0,0.000,0),(197,5207,'TOMÀQUET CHERRY ZEBRA','TOMÀQUET CHERRY ZEBRA','','','','','','',105,0,18,1,0,0,0,'',12.99,0,0.00,1,207,1,1,'',0,0,0.000,'',0,'','-',0,'V_107.png','','',0,0.000,0),(198,5208,'ALBERGÍNIA CAMP (EG)','ALBERGÍNIA CAMP (EG)','','','','','','',107,0,11,1,0,0,0,'',12.99,0,0.00,1,208,1,1,'',0,0,0.000,'',0,'','-',0,'V_013.png','','',0,0.000,0),(199,5209,'ALBERGÍNIA EXTRA','ALBERGÍNIA EXTRA','','','','','','',107,0,12,1,0,0,0,'',12.99,0,0.00,1,209,1,1,'',0,0,0.000,'',0,'','-',0,'V_014.png','','',0,0.000,0),(200,5210,'CARBASSÓ CAMP (EG)','CARBASSÓ CAMP (EG)','','','','','','',107,0,13,1,0,0,0,'',12.99,0,0.00,1,210,1,1,'',0,0,0.000,'',0,'','-',0,'V_022.png','','',0,0.000,0),(201,5211,'TOMÀQUET HORTA','TOMÀQUET HORTA','','','','','','',105,0,19,1,0,0,0,'',12.99,0,0.00,1,211,1,1,'',0,0,0.000,'',0,'','-',0,'V_113.png','','',0,0.000,0),(202,5212,'PEBROT VERD CAMP','PEBROT VERD CAMP','','','','','','',107,0,14,1,0,0,0,'',12.99,0,0.00,1,212,1,1,'',0,0,0.000,'',0,'','-',0,'V_082.png','','',0,0.000,0),(203,5213,'TOMÀQUET ATIGRAT','TOMÀQUET ATIGRAT','','','','','','',105,0,20,1,0,0,0,'',12.99,0,0.00,1,213,1,1,'',0,0,0.000,'',0,'','-',0,'V_110.png','','',0,0.000,0),(204,5214,'PEBROT VERMELL CAMP (EG)','PEBROT VERMELL CAMP (EG)','','','','','','',107,0,15,1,0,0,0,'',12.99,0,0.00,1,214,1,1,'',0,0,0.000,'',0,'','-',0,'V_084.png','','',0,0.000,0),(205,5215,'TOMÀQUET AMANIDA (EG)','TOMÀQUET AMANIDA (EG)','','','','','','',105,0,21,1,0,0,0,'',12.99,0,0.00,1,215,1,1,'',0,0,0.000,'',0,'','-',0,'V_102.png','','',0,0.000,0),(206,5216,'TOMÀQUET SABOR IRIS','TOMÀQUET SABOR IRIS','','','','','','',105,0,22,1,0,0,0,'',12.99,0,0.00,1,216,1,1,'',0,0,0.000,'',0,'','-',0,'V_107.png','','',0,0.000,0),(207,5217,'PEBROT ITALIÀ','PEBROT ITALIÀ','','','','','','',107,0,16,1,0,0,0,'',12.99,0,0.00,1,217,1,1,'',0,0,0.000,'',0,'','-',0,'V_086.png','','',0,0.000,0),(208,5218,'PEBROT ITALIÀ CAMP','PEBROT ITALIÀ CAMP','','','','','','',107,0,17,1,0,0,0,'',12.99,0,0.00,1,218,1,1,'',0,0,0.000,'',0,'','-',0,'V_082.png','','',0,0.000,0),(209,5219,'COGOMBRE ESPANYOL CAMP','COGOMBRE ESPANYOL CAMP','','','','','','',107,0,18,1,0,0,0,'',12.99,0,0.00,1,219,1,1,'',0,0,0.000,'',0,'','-',0,'V_075.png','','',0,0.000,0),(210,5220,'TOMÀQUET TOMATXOC','TOMÀQUET TOMATXOC','','','','','','',105,0,23,1,0,0,0,'',12.99,0,0.00,1,220,1,1,'',0,0,0.000,'',0,'','-',0,'V_111.png','','',0,0.000,0),(211,5221,'PÈSOLS','PÈSOLS','','','','','','',107,0,19,1,0,0,0,'',12.99,0,0.00,1,221,1,1,'',0,0,0.000,'',0,'','-',0,'V_051.png','','',0,0.000,0),(212,5222,'PASTANAGA GRAN (RC)','PASTANAGA GRAN (RC)','','','','','','',107,0,20,1,0,0,0,'',12.99,0,0.00,1,222,1,1,'',0,0,0.000,'',0,'','-',0,'V_116.png','','',0,0.000,0),(213,5223,'BOLET TROMPETA DE LA MORT','BOLET TROMPETA DE LA MORT','','','','','','',110,0,6,1,0,0,0,'',12.99,0,0.00,1,223,1,1,'',0,0,0.000,'',0,'','-',0,'V_016.png','','',0,0.000,0),(214,5224,'FAVES GRANELL','FAVES GRANELL','','','','','','',107,0,21,1,0,0,0,'',12.99,0,0.00,1,224,1,1,'',0,0,0.000,'',0,'','-',0,'V_101.png','','',0,0.000,0),(215,5225,'ALBERGÍNIA NEGRA','ALBERGÍNIA NEGRA','','','','','','',107,0,22,1,0,0,0,'',12.99,0,0.00,1,225,1,1,'',0,0,0.000,'',0,'','-',0,'V_012.png','','',0,0.000,0),(216,5226,'COGOMBRE ALMERIA','COGOMBRE ALMERIA','','','','','','',107,0,23,1,0,0,0,'',12.99,0,0.00,1,226,1,1,'',0,0,0.000,'',0,'','-',0,'V_076.png','','',0,0.000,0),(217,5227,'TOMÀQUET RAÏM','TOMÀQUET RAÏM','','','','','','',105,0,24,1,0,0,0,'',12.99,0,0.00,1,227,1,1,'',0,0,0.000,'',0,'','-',0,'V_108.png','','',0,0.000,0),(218,5228,'PEBROT VERMELL GRAN','PEBROT VERMELL GRAN','','','','','','',107,0,24,1,0,0,0,'',12.99,0,0.00,1,228,1,1,'',0,0,0.000,'',0,'','-',0,'V_088.png','','',0,0.000,0),(219,5230,'CARBASSÓ BLANC','CARBASSÓ BLANC','','','','','','',107,0,25,1,0,0,0,'',12.99,0,0.00,1,230,1,1,'',0,0,0.000,'',0,'','-',0,'V_023.png','','',0,0.000,0),(220,5231,'LLENGUA BOU (BOLET)','LLENGUA BOU (BOLET)','','','','','','',110,0,7,1,0,0,0,'',12.99,0,0.00,1,231,1,1,'',0,0,0.000,'',0,'','-',0,'V_016.png','','',0,0.000,0),(221,5232,'COGOMBRE NEGRE','COGOMBRE NEGRE','','','','','','',107,0,26,1,0,0,0,'',12.99,0,0.00,1,232,1,1,'',0,0,0.000,'',0,'','-',0,'V_075.png','','',0,0.000,0),(222,5233,'CARBASSÓ SAFATA 5U','CARBASSÓ SAFATA 5U','','','','','','',107,0,27,1,0,0,0,'',12.99,0,0.00,0,233,1,1,'',0,0,0.000,'',0,'','-',0,'V_025.png','','',0,0.000,0),(223,5234,'COGOMBRE ESPANYOL','COGOMBRE ESPANYOL','','','','','','',107,0,28,1,0,0,0,'',12.99,0,0.00,1,234,1,1,'',0,0,0.000,'',0,'','-',0,'V_076.png','','',0,0.000,0),(224,5235,'CARBASSÓ GRANELL','CARBASSÓ GRANELL','','','','','','',107,0,29,1,0,0,0,'',12.99,0,0.00,1,235,1,1,'',0,0,0.000,'',0,'','-',0,'V_025.png','','',0,0.000,0),(225,5236,'PEBROT VERMELL SAFATA','PEBROT VERMELL SAFATA','','','','','','',107,0,30,1,0,0,0,'',12.99,0,0.00,0,236,1,1,'',0,0,0.000,'',0,'','-',0,'V_084.png','','',0,0.000,0),(226,5237,'TOMÀQUET BRANCA SAFATA','TOMÀQUET BRANCA SAFATA','','','','','','',105,0,25,1,0,0,0,'',12.99,0,0.00,0,237,1,1,'',0,0,0.000,'',0,'','-',0,'V_113.png','','',0,0.000,0),(227,5238,'ALBERGÍNIES SAFATA','ALBERGÍNIES SAFATA','','','','','','',107,0,31,1,0,0,0,'',12.99,0,0.00,0,238,1,1,'',0,0,0.000,'',0,'','-',0,'V_013.png','','',0,0.000,0),(229,5239,'prova','prova','','','','','','',111,0,4,1,0,0,0,'',5.00,0,0.00,0,0,0,1,'',0,0,0.000,'',0,'','-',0,'','','',0,0.000,0);
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
  UNIQUE KEY `ParamKey` (`ParamKey`)
) ENGINE=InnoDB AUTO_INCREMENT=370 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localcfgscale`
--

LOCK TABLES `localcfgscale` WRITE;
/*!40000 ALTER TABLE `localcfgscale` DISABLE KEYS */;
INSERT INTO `localcfgscale` VALUES (8,'TOUCH/loader','1'),(10,'TOUCH/slave-icon','slave'),(24,'RESOURCES/user-icons-path','../share/pcscale/resources/icons/item'),(28,'NETWORK/scale-number','1'),(31,'NETWORK/scale-counter','1'),(80,'TOUCH/vendor-icon','vnd1'),(81,'TOUCH/device-icon','colors'),(87,'TOUCH/scale-icon','scale'),(241,'CORE/mup-group','1'),(242,'CORE/plu-board-group','1'),(243,'local-ticket/StartNumber','0'),(244,'local-ticket/MaximumNumber','0'),(289,'TOUCH/items-dnd','1'),(321,'TOUCH/GROUPS/scroll/default','board3'),(322,'TOUCH/GROUPS/VendorsGroup/default','vendors1'),(323,'sw-full-version','2.10 - Methone. Build id: 20160322 - 2131'),(327,'serial-number','863505'),(328,'TOUCH/name','Board-15Familias-116TeclasBlancas-v3-TicketTrasero'),(329,'PRINTER/lbl-format','J60X57'),(330,'PRINTER/prn-format','estandar+logo'),(332,'simplified-invoice-serial',''),(335,'CORE/key-groups-linked','0'),(340,'CORE/menu-on-modeform','1'),(342,'TOUCH/GROUPS/menu/default','menu01');
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
INSERT INTO `localcountries` VALUES (1,4,'AF','Afganistan'),(2,248,'AX','Illes Aland'),(3,8,'AL','Albània'),(4,12,'DZ','Algèria'),(5,16,'AS','Samoa Americana'),(6,20,'AD','Andorra'),(7,24,'AO','Angola'),(8,660,'AI','Anguilla'),(9,10,'AQ','Antàrtida'),(10,28,'AG','Antigua i Barbuda'),(11,32,'AR','Argentina'),(12,51,'AM','Armènia'),(13,533,'AW','Aruba'),(14,36,'AU','Austràlia'),(15,40,'AT','Àustria'),(16,31,'AZ','Azerbaitjan'),(17,44,'BS','Bahames'),(18,48,'BH','Bahrain'),(19,50,'BD','Bangla Desh'),(20,52,'BB','Barbados'),(21,112,'BY','Bielorússia'),(22,56,'BE','Bèlgica'),(23,84,'BZ','Belize'),(24,204,'BJ','Benín'),(25,60,'BM','Bermuda'),(26,64,'BT','Bhutan'),(27,68,'BO','Bolívia, Estat Plurinacional de'),(28,535,'BQ','Bonaire, Saint Eustatius and Saba'),(29,70,'BA','Bòsnia i Hercegovina'),(30,72,'BW','Botswana'),(31,74,'BV','Illa Bouvet'),(32,76,'BR','Brasil'),(33,86,'IO','Territori Britànic de l\'Oceà Índic'),(34,96,'BN','Brunei (Negara Brunei Darussalam)'),(35,100,'BG','Bulgària'),(36,854,'BF','Burkina Faso'),(37,108,'BI','Burundi'),(38,116,'KH','Cambodja'),(39,120,'CM','Camerun'),(40,124,'CA','Canadà'),(41,132,'CV','Cap Verd'),(42,136,'KY','Illes Caiman'),(43,140,'CF','República Centreafricana'),(44,148,'TD','Txad'),(45,152,'CL','Xile'),(46,156,'CN','Xina'),(47,162,'CX','Illa Christmas'),(48,166,'CC','Illes Cocos (Keeling)'),(49,170,'CO','Colòmbia'),(50,174,'KM','Comores'),(51,178,'CG','Congo'),(52,180,'CD','Congo, La República Democràtica del'),(53,184,'CK','Illes Cook'),(54,188,'CR','Costa Rica'),(55,384,'CI','Costa de Vori'),(56,191,'HR','Croàcia'),(57,192,'CU','Cuba'),(58,531,'CW','Curaçao'),(59,196,'CY','Xipre'),(60,203,'CZ','República Txeca'),(61,208,'DK','Dinamarca'),(62,262,'DJ','Djibouti'),(63,212,'DM','Dominica'),(64,214,'DO','República Dominicana'),(65,218,'EC','Equador'),(66,818,'EG','Egipte'),(67,222,'SV','El Salvador'),(68,226,'GQ','Guinea Equatorial'),(69,232,'ER','Eritrea'),(70,233,'EE','Estònia'),(71,231,'ET','Etiòpia'),(72,238,'FK','Illes Malvines (Falkland)'),(73,234,'FO','Illes Fèroe'),(74,242,'FJ','Fiji'),(75,246,'FI','Finlàndia'),(76,250,'FR','França'),(77,254,'GF','Guaiana Francesa'),(78,258,'PF','Polinèsia francesa'),(79,260,'TF','Territoris Francesos del Sud'),(80,266,'GA','Gabon'),(81,270,'GM','Gàmbia'),(82,268,'GE','Geòrgia'),(83,276,'DE','Alemanya'),(84,288,'GH','Ghana'),(85,292,'GI','Gibraltar'),(86,300,'GR','Grècia'),(87,304,'GL','Groenlàndia'),(88,308,'GD','Grenada'),(89,312,'GP','Guadalupe'),(90,316,'GU','Guam'),(91,320,'GT','Guatemala'),(92,831,'GG','Guernsey'),(93,324,'GN','Guinea'),(94,624,'GW','Guinea Bissau'),(95,328,'GY','Guyana'),(96,332,'HT','Haití'),(97,334,'HM','Illa Heard i Illes McDonald'),(98,336,'VA','Santa Seu (Estat del Vaticà)'),(99,340,'HN','Hondures'),(100,344,'HK','Hong Kong'),(101,348,'HU','Hongria'),(102,352,'IS','Islàndia'),(103,356,'IN','Índia'),(104,360,'ID','Indonèsia'),(105,364,'IR','Iran'),(106,368,'IQ','Iraq'),(107,372,'IE','Irlanda'),(108,833,'IM','Illa de Man'),(109,376,'IL','Israel'),(110,380,'IT','Itàlia'),(111,388,'JM','Jamaica'),(112,392,'JP','Japó'),(113,832,'JE','Jersey'),(114,400,'JO','Jordània'),(115,398,'KZ','Kazakhstan'),(116,404,'KE','Kenya'),(117,296,'KI','Kiribati'),(118,408,'KP','Corea del Nord'),(119,410,'KR','Corea del Sud'),(120,414,'KW','Kuwait'),(121,417,'KG','Kirguizistan'),(122,418,'LA','República Democràtica Popular de Laos'),(123,428,'LV','Letònia'),(124,422,'LB','Líban'),(125,426,'LS','Lesotho'),(126,430,'LR','Libèria'),(127,434,'LY','Libyan Arab Jamahiriya'),(128,438,'LI','Liechtenstein'),(129,440,'LT','Lituània'),(130,442,'LU','Luxemburg'),(131,446,'MO','Macau'),(132,807,'MK','Macedònia, República de'),(133,450,'MG','Madagascar'),(134,454,'MW','Malawi'),(135,458,'MY','Malàisia'),(136,462,'MV','Maldives'),(137,466,'ML','Mali'),(138,470,'MT','Malta'),(139,584,'MH','Illes Marshall'),(140,474,'MQ','Martinica'),(141,478,'MR','Mauritània'),(142,480,'MU','Maurici'),(143,175,'YT','Mayotte'),(144,484,'MX','Mèxic'),(145,583,'FM','Micronèsia, Estats Federats de'),(146,498,'MD','Moldàvia, República de'),(147,492,'MC','Mònaco'),(148,496,'MN','Mongòlia'),(149,499,'ME','Montenegro'),(150,500,'MS','Montserrat'),(151,504,'MA','Marroc'),(152,508,'MZ','Moçambic'),(153,104,'MM','Myanmar'),(154,516,'NA','Namíbia'),(155,520,'NR','Nauru'),(156,524,'NP','Nepal'),(157,528,'NL','Països Baixos'),(158,540,'NC','Nova Caledònia'),(159,554,'NZ','Nova Zelanda'),(160,558,'NI','Nicaragua'),(161,562,'NE','Níger'),(162,566,'NG','Nigèria'),(163,570,'NU','Niue'),(164,574,'NF','Illa Norfolk'),(165,580,'MP','Illes Marianes del Nord'),(166,578,'NO','Noruega'),(167,512,'OM','Oman'),(168,586,'PK','Pakistan'),(169,585,'PW','Palau'),(170,275,'PS','Territori Palestí, Ocupat'),(171,591,'PA','Panamà'),(172,598,'PG','Papua Nova Guinea'),(173,600,'PY','Paraguai'),(174,604,'PE','Perú'),(175,608,'PH','Filipines'),(176,612,'PN','Pitcairn'),(177,616,'PL','Polònia'),(178,620,'PT','Portugal'),(179,630,'PR','Puerto Rico'),(180,634,'QA','Qatar'),(181,638,'RE','Reunion'),(182,642,'RO','Romania'),(183,643,'RU','Federació Russa'),(184,646,'RW','Rwanda'),(185,652,'BL','Saint Barthélemy'),(186,654,'SH','Saint Helena, Ascension i Tristan da Cunha'),(187,659,'KN','Saint Christopher i Nevis'),(188,662,'LC','Saint Lucia'),(189,663,'MF','Saint Martin (part francesa)'),(190,666,'PM','Saint Pierre i Miquelon'),(191,670,'VC','Saint Vincent i les Grenadines'),(192,882,'WS','Samoa'),(193,674,'SM','San Marino'),(194,678,'ST','Sao Tome i Príncipe'),(195,682,'SA','Aràbia Saudita'),(196,686,'SN','Senegal'),(197,688,'RS','Sèrbia'),(198,690,'SC','Seychelles'),(199,694,'SL','Sierra Leone'),(200,702,'SG','Singapur'),(201,534,'SX','Sint Maarten'),(202,703,'SK','Eslovàquia'),(203,705,'SI','Eslovènia'),(204,90,'SB','Illes Salomó'),(205,706,'SO','Somàlia'),(206,710,'ZA','Sudàfrica'),(207,239,'GS','Illes Geòrgia del Sud i Sandwich del Sud'),(208,724,'ES','Espanya'),(209,144,'LK','Sri Lanka'),(210,736,'SD','Sudan'),(211,740,'SR','Surinam'),(212,744,'SJ','Svalbard i Jan Mayen'),(213,748,'SZ','Swazilàndia'),(214,752,'SE','Suècia'),(215,756,'CH','Suïssa'),(216,760,'SY','República Àrab Síria'),(217,158,'TW','Taiwan, Província de Xina'),(218,762,'TJ','Tadjikistan'),(219,834,'TZ','Tanzània, República Unida de'),(220,764,'TH','Tailàndia'),(221,626,'TL','Timor Oriental'),(222,768,'TG','Togo'),(223,772,'TK','Tokelau'),(224,776,'TO','Tonga'),(225,780,'TT','Trinitat i Tobago'),(226,788,'TN','Tunísia'),(227,792,'TR','Turquia'),(228,795,'TM','Turkmenistan'),(229,796,'TC','Illes Turks i Caicos'),(230,798,'TV','Tuvalu'),(231,800,'UG','Uganda'),(232,804,'UA','Ucraïna'),(233,784,'AE','Unió dels Emirats Àrabs'),(234,826,'GB','Regne Unit'),(235,840,'US','Estats Units'),(236,581,'UM','Illes Perifèriques Menors dels EUA'),(237,858,'UY','Uruguai'),(238,860,'UZ','Uzbekistan'),(239,548,'VU','Vanuatu'),(240,862,'VE','Venezuela, Bolivarian republic of'),(241,704,'VN','Vietnam'),(242,92,'VG','Illes Verges, Britàniques'),(243,850,'VI','Illes Verges, EUA'),(244,876,'WF','Wallis i Futuna'),(245,732,'EH','Sàhara Occidental'),(246,887,'YE','Iemen'),(247,894,'ZM','Zàmbia'),(248,716,'ZW','Zimbabwe');
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
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localerasetickets`
--

LOCK TABLES `localerasetickets` WRITE;
/*!40000 ALTER TABLE `localerasetickets` DISABLE KEYS */;
INSERT INTO `localerasetickets` VALUES (1,'T','2016-03-29 16:17:10'),(3,'P','2016-03-29 16:17:16'),(4,'S','2016-03-29 16:17:20');
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
  KEY `Invoice` (`Invoice`),
  KEY `IdNetHticket` (`IdNetHticket`),
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
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localnumtickets`
--

LOCK TABLES `localnumtickets` WRITE;
/*!40000 ALTER TABLE `localnumtickets` DISABLE KEYS */;
INSERT INTO `localnumtickets` VALUES (1,'I','1400101',8);
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
  `Version` smallint(6) DEFAULT '23',
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
INSERT INTO `localstatus` VALUES (1,70,'2.10.2131','ca_ES',0,0,0,0,0,0,0,0,0,0);
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
INSERT INTO `localvendorkeys` VALUES (1,1,101),(1,2,102),(1,3,103),(1,4,104),(1,5,105),(1,6,106),(1,7,107),(1,8,108),(1,9,109),(1,10,110),(1,11,111),(1,12,112),(1,21,121),(1,22,122);
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
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localvendorregister`
--

LOCK TABLES `localvendorregister` WRITE;
/*!40000 ALTER TABLE `localvendorregister` DISABLE KEYS */;
INSERT INTO `localvendorregister` VALUES (1,101,'2016-02-22 09:07:09',0,'2016-02-23 00:00:00',0);
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
) ENGINE=InnoDB AUTO_INCREMENT=641 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `localworkingmodes`
--

LOCK TABLES `localworkingmodes` WRITE;
/*!40000 ALTER TABLE `localworkingmodes` DISABLE KEYS */;
INSERT INTO `localworkingmodes` VALUES (3,'00 sales','based','sales'),(8,'01 packet','based','packet'),(9,'01 packet','description','Envasado'),(10,'01 packet','keyboard','packet'),(11,'01 packet','user','admin'),(12,'01 packet','keyboard-loop','packet-loop'),(13,'01 packet','in-use','1'),(14,'02 selfservice','based','selfservice'),(20,'03 pos','based','sales'),(93,'01 packet','user-check-mode-form','OFF'),(94,'01 packet','user-check-mode','OFF'),(95,'01 packet','user-prog-prices','ON'),(96,'01 packet','sections',''),(97,'01 packet','packet-customer','ON'),(98,'01 packet','show-print-packet','ON'),(99,'01 packet','allow-mult-key','ON'),(100,'01 packet','packet-vendor','0'),(101,'01 packet','show-form-packet','0'),(102,'01 packet','packet-after-item','OFF'),(103,'01 packet','autocode','ON'),(104,'01 packet','pluprc-key','ON'),(105,'01 packet','constant-key','ON'),(106,'01 packet','change-price','ON'),(107,'01 packet','amount-zero','ON'),(108,'01 packet','negative-amount','ON'),(109,'01 packet','code-digits','6'),(110,'01 packet','item-access','0'),(111,'01 packet','mixed-allowed','ON'),(112,'01 packet','weightlow-tsec','0'),(113,'01 packet','weightlow-beep','ON'),(114,'01 packet','weightlow-zero','OFF'),(129,'00 sales','allow-reload','ON'),(148,'03 pos','user-check-mode-form','OFF'),(149,'03 pos','user-check-mode','OFF'),(150,'03 pos','user-prog-prices','ON'),(151,'03 pos','sections',''),(152,'03 pos','print-ticket','ON'),(153,'03 pos','print-label-total','OFF'),(154,'03 pos','print-label-line','OFF'),(155,'03 pos','close-with-line','OFF'),(156,'03 pos','total-discount','ON'),(157,'03 pos','sales-customer','ON'),(158,'03 pos','allow-orders','ON'),(159,'03 pos','payment','0'),(160,'03 pos','startlocal','0'),(161,'03 pos','allow-cancel','ON'),(162,'03 pos','allow-reload','ON'),(163,'03 pos','allow-amount0','ON'),(164,'03 pos','auto-fix','ON'),(165,'03 pos','times-print','250'),(167,'03 pos','autocode','OFF'),(168,'03 pos','pluprc-key','ON'),(169,'03 pos','constant-key','ON'),(170,'03 pos','change-price','ON'),(171,'03 pos','return-weight','ON'),(172,'03 pos','amount-zero','ON'),(173,'03 pos','negative-amount','ON'),(174,'03 pos','code-digits','6'),(175,'03 pos','item-access','0'),(176,'03 pos','pos-method','ON'),(177,'03 pos','clear-tare','OFF'),(178,'03 pos','weightlow-tsec','0'),(179,'03 pos','weightlow-beep','ON'),(180,'03 pos','weightlow-zero','OFF'),(191,'02 selfservice','description','Autoservicio'),(192,'02 selfservice','keyboard','packet'),(193,'02 selfservice','user','admin'),(194,'02 selfservice','keyboard-loop',''),(195,'02 selfservice','keyboard-vendors',''),(196,'02 selfservice','in-use','0'),(203,'00 sales','description','Balanza'),(204,'00 sales','keyboard','sales'),(205,'00 sales','user','admin'),(206,'00 sales','keyboard-loop',''),(207,'00 sales','keyboard-vendors',''),(208,'00 sales','in-use','1'),(213,'','working-mode','00 sales'),(262,'00 sales','user-check-mode-form','OFF'),(263,'00 sales','user-check-mode','OFF'),(264,'00 sales','user-prog-prices','ON'),(265,'00 sales','sections',''),(266,'00 sales','print-ticket','ON'),(267,'00 sales','print-label-total','OFF'),(268,'00 sales','print-label-line','OFF'),(269,'00 sales','close-with-line','OFF'),(270,'00 sales','total-discount','ON'),(271,'00 sales','sales-customer','ON'),(272,'00 sales','allow-orders','OFF'),(273,'00 sales','payment','2'),(274,'00 sales','startlocal','3'),(275,'00 sales','allow-cancel','ON'),(276,'00 sales','allow-amount0','ON'),(277,'00 sales','auto-fix','OFF'),(278,'00 sales','times-print','250'),(280,'00 sales','auto-vendor','ON'),(281,'00 sales','customer-care','OFF'),(282,'00 sales','park-ticket','ON'),(283,'00 sales','ticket-format',''),(284,'00 sales','label-format',''),(285,'00 sales','parking-format','parked'),(286,'00 sales','reload-type','48'),(287,'00 sales','reprint-type','240'),(288,'00 sales','allow-training','ON'),(289,'00 sales','autocode','ON'),(290,'00 sales','pluprc-key','ON'),(291,'00 sales','constant-key','ON'),(292,'00 sales','change-price','ON'),(293,'00 sales','return-weight','ON'),(294,'00 sales','amount-zero','ON'),(295,'00 sales','negative-amount','ON'),(296,'00 sales','code-digits','6'),(297,'00 sales','item-access','0'),(298,'00 sales','pos-method','OFF'),(299,'00 sales','clear-tare','ON'),(300,'00 sales','weightlow-tsec','0'),(301,'00 sales','weightlow-beep','ON'),(302,'00 sales','weightlow-zero','OFF'),(303,'03 pos','description','POS'),(304,'03 pos','keyboard','packet'),(305,'03 pos','user','admin'),(306,'03 pos','key-fnc-on-pressed','0'),(307,'03 pos','keyboard-loop',''),(308,'03 pos','keyboard-vendors',''),(309,'03 pos','PLUload-func','0'),(310,'03 pos','PLUload-mode','0'),(311,'03 pos','in-use','0'),(316,'00 sales','key-fnc-on-pressed','0'),(319,'00 sales','PLUload-func','2'),(320,'00 sales','PLUload-mode','1'),(334,'01 packet','key-fnc-on-pressed','0'),(336,'01 packet','keyboard-vendors',''),(337,'01 packet','PLUload-func','2'),(338,'01 packet','PLUload-mode','1'),(341,'00 sales','vendor-hide-code','OFF'),(342,'00 sales','vendor-summary','OFF'),(343,'00 sales','single-section','OFF'),(344,'00 sales','ean13-section-item','OFF'),(345,'00 sales','ticket-days','1'),(346,'00 sales','simplified-invoice','OFF'),(347,'00 sales','hold-plusec','ON'),(348,'00 sales','plukey-dnd','Full'),(390,'','user-check-menu-form','OFF'),(391,'00 sales','current-data-mode','0'),(420,'00 sales','user-pricerates','ON'),(456,'00 sales','pricerate','0'),(478,'02 selfservice','key-fnc-on-pressed','0'),(481,'02 selfservice','PLUload-func','2'),(482,'02 selfservice','PLUload-mode','0'),(493,'04 salesturn','based','sales'),(494,'04 salesturn','description','Venta+Turno'),(495,'04 salesturn','keyboard','sales_Turn'),(496,'04 salesturn','user','admin'),(497,'04 salesturn','key-fnc-on-pressed','0'),(498,'04 salesturn','PLUload-func','0'),(499,'04 salesturn','PLUload-mode','1'),(500,'04 salesturn','in-use','1'),(502,'04 salesturn','current-data-mode','0'),(507,'04 salesturn','keyboard-loop',''),(508,'04 salesturn','keyboard-vendors',''),(509,'01 packet','current-data-mode','0'),(511,'04 salesturn','user-check-mode-form','OFF'),(512,'04 salesturn','user-check-mode','OFF'),(513,'04 salesturn','user-prog-prices','ON'),(514,'04 salesturn','sections',''),(515,'04 salesturn','print-ticket','ON'),(516,'04 salesturn','print-label-total','OFF'),(517,'04 salesturn','print-label-line','OFF'),(518,'04 salesturn','close-with-line','OFF'),(519,'04 salesturn','total-discount','ON'),(520,'04 salesturn','sales-customer','ON'),(521,'04 salesturn','allow-orders','ON'),(522,'04 salesturn','payment','2'),(523,'04 salesturn','startlocal','0'),(524,'04 salesturn','allow-cancel','ON'),(525,'04 salesturn','allow-amount0','ON'),(526,'04 salesturn','auto-fix','OFF'),(527,'04 salesturn','times-print','250'),(528,'04 salesturn','pos-method','OFF'),(529,'04 salesturn','auto-vendor','ON'),(530,'04 salesturn','vendor-hide-code','OFF'),(531,'04 salesturn','vendor-summary','OFF'),(532,'04 salesturn','customer-care','OFF'),(533,'04 salesturn','park-ticket','ON'),(534,'04 salesturn','single-section','OFF'),(535,'04 salesturn','ean13-section-item','OFF'),(536,'04 salesturn','ticket-format',''),(537,'04 salesturn','label-format',''),(538,'04 salesturn','parking-format','parked'),(539,'04 salesturn','reload-type','1'),(540,'04 salesturn','reprint-type','1'),(541,'04 salesturn','ticket-days','1'),(542,'04 salesturn','allow-training','ON'),(543,'04 salesturn','simplified-invoice','ON'),(544,'04 salesturn','hold-plusec','ON'),(545,'04 salesturn','plukey-dnd','Full'),(546,'04 salesturn','user-pricerates','ON'),(547,'04 salesturn','autocode','ON'),(548,'04 salesturn','pluprc-key','ON'),(549,'04 salesturn','constant-key','ON'),(550,'04 salesturn','change-price','ON'),(551,'04 salesturn','return-weight','ON'),(552,'04 salesturn','amount-zero','OFF'),(553,'04 salesturn','negative-amount','ON'),(554,'04 salesturn','code-digits','6'),(555,'04 salesturn','item-access','0'),(556,'04 salesturn','clear-tare','ON'),(557,'04 salesturn','weightlow-tsec','0'),(558,'04 salesturn','weightlow-beep','ON'),(559,'04 salesturn','weightlow-zero','OFF'),(560,'04 salesturn','pricerate','0'),(562,'03 pos','reprint-ticket','ON'),(563,'03 pos','reprint-label-total','OFF'),(564,'00 sales','reprint-ticket','ON'),(565,'00 sales','reprint-label-total','OFF'),(566,'04 salesturn','reprint-ticket','ON'),(567,'04 salesturn','reprint-label-total','OFF');
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
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `ltickets`
--

LOCK TABLES `ltickets` WRITE;
/*!40000 ALTER TABLE `ltickets` DISABLE KEYS */;
INSERT INTO `ltickets` VALUES (8,8,'N',1,1,'-',5002,'TARONJA GRANELL RC',0,101,0,1,101,'2016-03-29 16:21:39','P',0,1,4.00,0.105,0.000,'',0.00,12.99,1.36,0,'','','0000-00-00',0.00,1,'2016-03-29 16:21:51',1,'T',0,'0000-00-00 00:00:00',0,''),(9,9,'N',1,1,'-',5003,'LLIMONES BIO',0,101,0,1,101,'2016-03-29 16:22:14','P',0,1,4.00,0.105,0.000,'',0.00,12.99,1.36,0,'','','0000-00-00',0.00,1,'2016-03-29 16:22:21',2,'T',0,'0000-00-00 00:00:00',0,''),(10,10,'N',1,1,'-',5109,'PRÉSSEC GROC',0,103,0,1,101,'2016-03-29 16:22:54','P',0,1,4.00,0.105,0.000,'',0.00,12.99,1.36,0,'','','0000-00-00',0.00,1,'2016-03-29 16:22:58',3,'T',0,'0000-00-00 00:00:00',0,''),(11,11,'N',1,1,'-',5020,'PERA DEGANA',0,102,0,2,101,'2016-03-29 16:23:14','P',0,1,4.00,0.100,0.000,'',0.00,12.99,1.30,0,'','','0000-00-00',0.00,2,'2016-03-29 16:23:23',4,'T',0,'0000-00-00 00:00:00',0,''),(12,12,'N',1,1,'-',5114,'PARAGUAIÀ LLEIDA',0,103,0,1,101,'2016-03-29 16:24:05','P',0,1,4.00,0.105,0.000,'',0.00,12.99,1.36,0,'','','0000-00-00',0.00,1,'2016-03-29 16:24:08',5,'T',0,'0000-00-00 00:00:00',0,''),(13,13,'L',1,1,'-',5225,'ALBERGÍNIA NEGRA',0,107,0,2,101,'2016-03-29 16:24:31','P',0,1,4.00,0.100,0.000,'',0.00,12.99,1.30,0,'','','0000-00-00',0.00,2,'2016-03-29 16:24:38',1,'T',0,'0000-00-00 00:00:00',0,''),(14,14,'L',1,0,'-',0,'',0,0,1,2,101,'2016-03-29 16:25:10','N',1,0,0.00,1.000,0.000,'',0.00,0.01,0.01,0,'','','0000-00-00',0.00,2,'2016-03-29 16:25:18',2,'T',0,'0000-00-00 00:00:00',0,''),(15,15,'N',1,1,'-',5157,'PINYA',0,104,0,2,101,'2016-03-29 16:25:40','P',0,1,4.00,0.100,0.000,'',0.00,12.99,1.30,0,'','','0000-00-00',0.00,2,'2016-03-29 16:25:47',6,'T',0,'0000-00-00 00:00:00',0,'');
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
INSERT INTO `netstatus` VALUES (1,'Servidora Taller','Servidora Taller 2016/03/29 09:47:56 1972341997',2335);
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
INSERT INTO `numberingtickets` VALUES (2,'P',1,0,1,99999,9999,1,'D',0,0,0,0),(3,'S',1,0,1,99999,9999,1,'D',0,0,0,0),(4,'C',1,0,1,99999,9999,0,'D',0,0,0,0),(5,'Z',1,0,1,9999,9999,0,'D',0,0,0,0),(6,'D',1,0,1,9999,9999,0,'D',0,0,0,0),(7,'F',1,0,1,9999,9999,0,'D',0,0,0,0),(8,'T',1,1,1,99999,99999,1,'D',0,0,0,0);
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
INSERT INTO `numinvoices` VALUES (1,99999,0,'2016/',0,'R');
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
) ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `numtickets`
--

LOCK TABLES `numtickets` WRITE;
/*!40000 ALTER TABLE `numtickets` DISABLE KEYS */;
INSERT INTO `numtickets` VALUES (4,'D','1',4),(5,'I','1600101',37),(8,'F','1',5),(10,'T','1',6);
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
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `paymethods`
--

LOCK TABLES `paymethods` WRITE;
/*!40000 ALTER TABLE `paymethods` DISABLE KEYS */;
INSERT INTO `paymethods` VALUES (1,0,'EFECTIU',1),(2,1,'TARJA',1),(3,2,'XEC',1),(4,3,'PAYMETHOD3',0),(5,4,'PAYMETHOD4',0),(6,5,'PAYMETHOD5',0),(7,6,'PAYMETHOD6',0);
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
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `paytickets`
--

LOCK TABLES `paytickets` WRITE;
/*!40000 ALTER TABLE `paytickets` DISABLE KEYS */;
INSERT INTO `paytickets` VALUES (8,8,0,1.36),(9,9,0,1.36),(10,10,0,1.36),(11,11,0,1.30),(12,12,0,1.36),(13,13,0,1.30),(14,14,0,0.01),(15,15,0,1.30);
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
) ENGINE=InnoDB AUTO_INCREMENT=51 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `plukeys`
--

LOCK TABLES `plukeys` WRITE;
/*!40000 ALTER TABLE `plukeys` DISABLE KEYS */;
INSERT INTO `plukeys` VALUES (1,0,1,1,5001),(2,0,1,2,5011),(3,0,1,3,5021),(4,0,1,4,5031),(5,0,1,5,5041),(6,0,1,6,5051),(7,0,1,7,5098),(8,0,1,8,5064),(9,0,1,9,5061),(10,0,1,10,5071),(11,0,1,11,5081),(12,0,1,12,5091),(13,0,1,13,5101),(14,0,1,14,5111),(15,0,1,15,5121),(16,0,1,16,5131),(17,0,1,17,5141),(18,0,1,18,5151),(19,0,1,19,5161),(20,0,1,20,5171),(21,0,1,21,5181),(22,0,1,22,5191),(23,0,1,23,5201),(24,0,1,24,5211),(25,0,1,25,5221),(26,0,1,26,5231),(27,0,1,27,5020),(28,0,1,28,5040),(29,0,1,29,5060),(30,0,1,30,5080),(31,0,1,31,5100),(32,0,1,32,5120),(33,0,1,33,5140),(34,0,1,34,5160),(35,0,1,35,5180),(37,0,1,36,5200),(38,0,1,37,5200),(39,0,1,38,5220),(40,0,1,39,5224),(41,0,1,40,5015),(43,0,1,41,5130),(44,0,1,42,5145),(45,0,1,43,5165),(46,0,1,44,5185),(47,0,1,45,5146),(48,0,1,46,5207),(49,0,1,47,5196),(50,0,1,48,5219);
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
  `SerialNumber` varchar(10) DEFAULT '',
  `System` varchar(32) DEFAULT '',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `MAC` (`MAC`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `scales`
--

LOCK TABLES `scales` WRITE;
/*!40000 ALTER TABLE `scales` DISABLE KEYS */;
INSERT INTO `scales` VALUES (6,'00:30:18:C0:37:9D','192.168.1.207',1,1,'','S','2.10.2131','863505','Ubuntu_precise'),(7,'00:30:18:C0:35:CE','192.168.1.203',2,1,'','N','2.10.2131','863244','Ubuntu_precise');
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
  `Version` smallint(6) DEFAULT '60',
  `AppVersion` varchar(16) DEFAULT '',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `status`
--

LOCK TABLES `status` WRITE;
/*!40000 ALTER TABLE `status` DISABLE KEYS */;
INSERT INTO `status` VALUES (1,70,'2.10.2131');
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
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totaldaily`
--

LOCK TABLES `totaldaily` WRITE;
/*!40000 ALTER TABLE `totaldaily` DISABLE KEYS */;
INSERT INTO `totaldaily` VALUES (8,'T','N','2016-03-29',0,1,0,0,6,6,8.04,0.00,0.620,0,0,0.00,0,0.00,2,2.60),(13,'T','L','2016-03-29',0,1,0,0,2,2,1.31,0.00,0.100,1,0,0.00,0,0.00,1,1.30);
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
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalpaytickets`
--

LOCK TABLES `totalpaytickets` WRITE;
/*!40000 ALTER TABLE `totalpaytickets` DISABLE KEYS */;
INSERT INTO `totalpaytickets` VALUES (8,'N','2016-03-29',0,8.04),(13,'L','2016-03-29',0,1.31);
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
) ENGINE=InnoDB AUTO_INCREMENT=50 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalscalefamilies`
--

LOCK TABLES `totalscalefamilies` WRITE;
/*!40000 ALTER TABLE `totalscalefamilies` DISABLE KEYS */;
INSERT INTO `totalscalefamilies` VALUES (1,'2016-02-01',1,1,101,1,0,0,0.000,9),(2,'2016-02-01',1,1,110,1,0,0,1.000,0),(3,'2016-02-01',1,1,102,1,0,0,1.195,0),(4,'2016-02-01',1,1,108,1,0,0,2.000,0),(5,'2016-02-01',1,1,0,1,0,0,4.000,3),(6,'2016-02-01',1,1,104,1,0,0,3.000,0),(7,'2016-02-01',1,1,106,1,0,0,3.000,3),(10,'2016-02-01',1,1,107,1,0,0,1.000,0),(11,'2016-02-01',3,3,0,1,0,0,2.000,4),(12,'2016-02-01',3,3,105,1,0,0,1.000,0),(13,'2016-02-01',3,3,108,1,0,0,1.000,0),(14,'2016-02-01',2,2,0,1,0,0,0.000,0),(15,'2016-02-01',2,2,101,1,0,0,0.000,2),(16,'2016-02-01',2,2,106,1,0,0,0.000,4),(21,'2016-02-01',3,3,104,1,0,0,1.000,0),(22,'2016-02-01',3,3,106,1,0,0,1.000,1),(24,'2016-02-01',2,2,102,1,0,0,1.000,0),(25,'2016-02-01',2,2,109,1,0,0,2.000,0),(27,'2016-02-01',3,3,103,1,0,0,0.000,0),(28,'2016-02-01',3,3,111,1,0,0,0.000,2),(30,'2016-02-01',1,1,103,1,0,0,1.000,0),(31,'2016-02-01',1,1,111,1,0,0,0.000,3),(34,'2016-02-01',2,2,104,1,0,0,0.000,0),(35,'2016-03-01',4,1,101,1,0,0,0.315,0),(36,'2016-03-01',5,2,102,1,0,0,0.310,0),(38,'2016-03-01',5,2,0,1,0,0,0.000,5),(44,'2016-03-01',0,1,103,1,0,0,0.210,0),(47,'2016-03-01',0,2,107,1,0,0,0.100,0),(49,'2016-03-01',0,2,104,1,0,0,0.100,0);
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
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `totalvat`
--

LOCK TABLES `totalvat` WRITE;
/*!40000 ALTER TABLE `totalvat` DISABLE KEYS */;
INSERT INTO `totalvat` VALUES (8,'N','2016-03-29',1,4.00,8.04,7.74,0.30),(13,'L','2016-03-29',1,4.00,1.30,1.25,0.05),(14,'L','2016-03-29',0,0.00,0.01,0.01,0.00);
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
INSERT INTO `trademark` VALUES (1,'','','','','','','','DEMO FRUITES I VERDURES','C\\FLORIDABLANCA 98 TDA.2','BARCELONA','08015','','TEL. 934 248 455','FAX. 934 240 265','','barnapes@barnapes.es','NIF. B60544244','www.barnapes.es','GRÀCIES','PER LA SEVA VISITA','','','','','','','','');
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
INSERT INTO `vats` VALUES (2,0,'Sin IVA',0.00),(3,1,'IVA 4 %',4.00),(4,2,'IVA 10 %',10.00),(5,3,'IVA 21 %',21.00),(6,4,'',0.00);
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
INSERT INTO `vendorkeys` VALUES (1,1,101),(1,2,102),(1,3,103),(1,4,104),(1,5,105),(1,6,106),(1,7,107),(1,8,108),(1,9,109),(1,10,110),(1,11,111),(1,12,112),(1,21,121),(1,22,122);
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
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
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
INSERT INTO `vendors` VALUES (101,'Cindy','',''),(102,'David','',''),(103,'Belinda','',''),(104,'Brad','',''),(105,'Catherine','',''),(106,'Colin','',''),(107,'Elsa','',''),(108,'George','',''),(109,'Jessica','',''),(110,'Eva','',''),(111,'Jack','',''),(112,'Keira','',''),(121,'Venedor 121','',''),(122,'Venedor 122','','');
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

-- Dump completed on 2016-03-29 16:42:05
