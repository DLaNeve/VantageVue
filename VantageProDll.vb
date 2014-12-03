Imports System.Runtime.InteropServices

Module VantageProDll
    Public Const FAHRENHEIT = 0
    Public Const CELSIUS = 1
    Public Const INCHES = 0
    Public Const MM = 1
    Public Const MB = 2
    Public Const HECTO_PASCAL = 3
    Public Const MPH = 0
    Public Const KNOTS = 1
    Public Const KPH = 2
    Public Const METERS_PER_SECOND = 3
    Public Const KM = 1
    Public Const MILES = 0
    Public Const FT = 1
    Public Const M = 0


    Public Const NOT_SET = -100
    Public Const COM_ERROR = -101
    Public Const MEMORY_ERROR = -102
    Public Const COM_OPEN = -103

    Public Const ENGLISH_10 = 0
    Public Const ENGLISH_100 = 1
    Public Const METRIC_5 = 2
    Public Const METRIC_1 = 3
    Public Const ENGLISH_OTHER = 4
    Public Const METRIC_OTHER = 5

    ' These are the indexes of the "system Timeouts" used in SetSystemTimeOuts()
    Public Const TO_STANDARD = 0
    Public Const TO_DUMP_AFTER = 1
    Public Const TO_MODEM = 2
    Public Const TO_LOOPBACK = 3
    Public Const TO_LOOP = 4
    Public Const TO_FLUSH = 5
    Public Const TO_DONE = 6
    Public Const TO_STANDARD_MODEM = 7
    Public Const TO_STANDARD_MONITOR = 8
    Public Const TO_STANDARD_MONITOR_MODEM = 9
    Public Const TO_AUTO_DETECT = 10

    Public Structure SensorImage
        Dim outsideTemp As Single
        Dim insideTemp As Single
        Dim insideHum As Integer
        Dim outsideHum As Int32
        Dim barometer As Single
        Dim windSpeed As Int32
        Dim windDirection As Int32
        Dim TotalRain As Single
    End Structure

    Public Structure DateTime
        Dim month As Int32
        Dim day As Int32
        Dim hour As Int32
        Dim min As Int32
    End Structure

    Structure DateTimeStamp
        Dim minute As Long
        Dim hour As Long
        Dim day As Long
        Dim month As Long
        Dim year As Long
    End Structure

    Structure TimeHourMin
        Dim hour As Long
        Dim min As Long
    End Structure


    Structure Options
        Dim UseDialogBox As Byte
        Dim ClearArchive As Byte
        Dim WriteToFile As Byte
        Dim AppendToFile As Byte
    End Structure

    Structure WeatherUnits
        Dim TempUnit As Byte
        Dim RainUnit As Byte
        Dim BaromUnit As Byte
        Dim WindUnit As Byte
        Dim elevUnit As Byte
    End Structure


    Structure WeatherRecordStruct
        Dim year As Int32
        Dim month As Byte
        Dim day As Byte
        Dim packedTime As Int32
        Dim dateStr As String
        Dim timeStr As String
        Dim heatIndex As Single
        Dim windChill As Single
        Dim hiOutsideTemp As Single
        Dim lowOutsideTemp As Single
        Dim dewPolong As Single
        Dim windSpeed As Single
        Dim windDirection As Single
        Dim windDirectionStr As String
        Dim hiWindSpeed As Single
        Dim rain As Single
        Dim barometer As Single
        Dim insideTemp As Single
        Dim outsideTemp As Single
        Dim insideHum As Single
        Dim outsideHum As Single
        Dim archivePeriod As Int32
    End Structure


    Structure CurrentVantageData
        Dim outsideTemp As Int32
        Dim insideTemp As Int32
        Dim insideHum As Int32
        Dim outsideHum As Int32
        Dim barometer As Int32
        Dim windSpeed As Int32
        Dim wind10Avg As Int32
        Dim windDirection As Int32
        Dim dayRain As Int32
        Dim monthRain As Int32
        Dim stormRain As Int32
        Dim dayEt As Int32
        Dim monthEt As Int32
        Dim stormEt As Int32
        Dim solarRad As Int32
        Dim uv As Byte
        Dim forecastIcons As Byte
        Dim forecastRule As Byte
        Dim insideAlarms As Byte
        Dim rainAlarms As Byte
        Dim outsideAlarms As Byte
        Dim sunriseTime As Int32
        Dim sunsetTime As Int32
        Dim txBatteryState As Byte
        Dim consoleBattery As Int32
    End Structure

    Class VantageAlarmDataTag
        Dim barRiseA As Byte
        Dim barFallA As Byte
        Dim timeA As Int32
        Dim timeCompA As Int32
        Dim lowInTempA As Byte
        Dim hiInTempA As Byte
        Dim lowOutTempA As Byte
        Dim hiOutTempA As Byte
        Dim lowExTempA(15) As Byte
        Dim hiExTempA(15) As Byte
        Dim lowInHumA As Byte
        Dim hiInHumA As Byte
        Dim lowExHumA(8) As Byte
        Dim hiExHumA(8) As Byte
        Dim lowDewPtA As Byte
        Dim hiDewPtA As Byte
        Dim lowChillA As Byte
        Dim hiHeatA As Byte
        Dim hiTHSWA As Byte
        Dim hiSpeedA As Byte
        Dim hi10MinSpeedA As Byte
        Dim hiUvA As Byte
        Dim hiUvDose As Byte
        Dim lowSoilA(4) As Byte
        Dim hiSoilA(4) As Byte
        Dim lowLeafA(4) As Byte
        Dim hiLeafA(4) As Byte
        Dim hiSolarA As Int32
        Dim hiRRateA As Int32
        Dim hiFFloodA As Int32
        Dim hiRDayA As Int32
        Dim hiRStormA As Int32
    End Class

    Class VantageAlarmData
        Dim rawData(93) As Byte
        Dim alarms As VantageAlarmDataTag
    End Class

    Structure VantageHiLowBlock
        Dim lowOutsideTemp As Int32
        Dim hiOutsideTemp As Int32
        Dim lowOutsideTempTime As DateTime
        Dim hiOutsideTempTime As DateTime

        Dim hiInsideTemp As Int32
        Dim lowInsideTemp As Int32
        Dim hiInsideTempTime As DateTime
        Dim lowInsideTempTime As DateTime

        Dim lowOutsideHum As Int32
        Dim hiOutsideHum As Int32
        Dim lowInsideHumTime As DateTime
        Dim hiOutsideHumTime As DateTime

        Dim hiInsideHum As Int32
        Dim lowInsideHum As Int32
        Dim hiInsideHumTime As DateTime
        Dim lowOutsideHumTime As DateTime

        Dim lowDewPolong As Single                       ' In F
        Dim hiDewPolong As Single                      'In F
        Dim lowDewPointTime As DateTime
        Dim hiDewPointTime As DateTime

        Dim hiTempHumIndex As Single
        Dim hiTempHumIndexTime As DateTime

        Dim hiWindSpeed As Int32                     'Stored in Mph.
        Dim hiWindSpeedTime As DateTime

        Dim lowWindChill As Single
        Dim lowWindChillTime As DateTime
        Dim lowBar As Int32
        Dim lowBarTime As DateTime

        Dim hiBar As Int32
        Dim hiBarTime As DateTime

        Dim hiRate As Single
        Dim hiRateTime As DateTime

        Dim hiUV As Single
        Dim hiUVTime As DateTime

        Dim hiSolar As Int32
        Dim hiSolarTime As DateTime
    End Structure

    Class VantageCalDataTag
        Dim tempIn As Int32
        Dim tempOut As Int32
        Dim tempEx(15) As Int32
        Dim humIn As Byte
        Dim humEx(8) As Byte
    End Class


    Class vantageCalData
        Dim rawData(43) As Byte
        Dim values As VantageCalDataTag
    End Class

    Structure CurrentVantageCalibration
        Dim tempIn As Single
        Dim tempOut As Single
        Dim humIn As Byte
        Dim humOut As Byte

        ' calibration offsets
        Dim tempInOffset As Single
        Dim tempOutOffset As Single
        Dim humInOffset As Byte
        Dim humOutOffset As Byte
    End Structure

    'WeatherRecordStructEx
    'Holds a data record from the VantagePro archive memory. The data is stored 
    'in the units currently selected in the DLL. This structure contains more 
    'data fields than the WeatherRecordStruct used by previous versions of the 
    'DLL and is filled in by GetArchiveRecordEx_V. 
    Class WeatherRecordStructEx
        Dim year As Int16
        Dim month As Byte
        Dim day As Byte
        Dim packedTime As Int16
        Dim dateStr(16) As Byte
        Dim timeStr(16) As Byte
        Dim archivePeriod As Int16

        Dim outsideTemp As Single
        Dim hiOutsideTemp As Single
        Dim lowOutsideTemp As Single
        Dim insideTemp As Single

        Dim barometer As Single
        Dim barometerTrend As Int16

        Dim outsideHum As Single
        Dim insideHum As Single

        Dim rain As Single
        Dim hiRainRate As Single

        Dim windSpeed As Single
        Dim hiWindSpeed As Single
        Dim windDirection As Int16
        Dim windDirectionStr(5) As Byte
        Dim hiWindDirection As Int16
        Dim hiWindDirectionStr(5) As Byte

        Dim numWindSamples As Int16
        Dim numExpectedSamples As Int16

        Dim solarRad As Int16
        Dim hiSolarRad As Int16
        Dim UV As Single
        Dim hiUV As Single

        Dim et As Single

        Dim extraTemp(3) As Single
        Dim extraHum(2) As Single
        Dim soilTemp(4) As Single
        Dim leafTemp(2) As Single

        Dim soilMoisture(4) As Single
        Dim leafWetness(2) As Single

        Dim heatIndex As Single
        Dim THWIndex As Single
        Dim THSWIndex As Single
        Dim windChill As Single
        Dim dewPoint As Single

        Dim insideHeatIndex As Single
        Dim insideDewPoint As Single
    End Class

    '   ReceptionStats
    ' Holds the reception statistics for the ISS or wireless anemometer station on 
    '  the VantagePro console since midnight or since they were cleared manualy on 
    'the console. The data is filled in by GetReceptionData_V. 
    Structure ReceptionStats
        Dim totalPacketsReceived As Long
        Dim totalPacketsMissed As Long
        Dim numberOfResynchs As Long
        Dim maxInARow As Long
        Dim numCRCerrors As Long
    End Structure
    '   LatLonValue
    'Holds a latitude or longitude value. They can be expressed either as a floating 
    'point number that holds degrees and fractions or as integer degrees, minutes, 
    'and seconds. Both set of data fields are filled in when reading a value from 
    'the VantagePro. The "bUseFranctionalDegrees" field is used to select which set 
    'of data the DLL should use when writing the data to the weather station (0 = use 
    'the integer degrees/minutes/seconds fields; 1 = use the fractionalDegrees field). 

    'Positive degree values are used for North latitude and East longitude. Negitive 
    'degrees are used for South latitude and West longitude. When using the integer 
    'degrees/minutes/seconds fields, only the degrees value is negative. The minutes 
    'and seconds fields are always positive numbers between 0 and 60. 

    'The data is filled in by GetVantageLat_V and GetVantageLon_V. The data structure 
    'is used to write new values to the VantagePro by SetVantageLat_V and SetVantageLon_V.
    Structure LatLonValue
        Dim bUseFractionalDegrees As Int16
        Dim fractionalDegrees As Int16
        Dim bNegativeDegrees As Int16
        Dim degrees As Int16
        Dim minutes As Int16
        Dim seconds As Int16
    End Structure

    '   TxConfiguration
    'Holds the transmiter configuration data read from the VantagePro console, or to 
    'be written to it. The data fields are filled in by GetVantageTxConfig_V and 
    'written to the weather station by SetVantageTxConfig_V. 

    'For each of the 8 transmitter ID's the corresponding txType entry indicates the 
    'selected Weather Transmitter Type, and the repeater entry indicates the selected 
    'VantagePro 2 repeater. For the repeater entry, use 0 for no repeater (or for 
    'VantagePro 1 systems), or a value from 1 to 8 to select repeater A through H. 

    'Note: The txType values used by the DLL do not match the ones specified in the 
    'Vantage Programmers reference. The DLL will determine the weather station 
    'firmware version and write the correct values to the station. 
    Class TxConfiguration
        Dim txType(8) As Int16
        Dim repeater(8) As Int16
    End Class

    ' BarCalData
    'Holds information about the current sea-level correction for the barometer 
    'sensor on the VantagePro console. The data values are filled in by 
    'GetBarometerData_V. 

    'Barometer, temperature, and elevation values are given in the current DLL units. 
    'The correntBarometer value is the most recently measured barometer reading (normaly 
    'updated every 15 minutes) corrected to sea-level. To determine the raw reading, 
    'subtract the barCalibrationOffset from it and divide the result by the 
    'barCalibrationRatio. 

    Structure BarCalData
        Dim currentBarometer As Single
        Dim elevation As Single
        Dim dewPoint As Single
        Dim virtualTemp As Single
        Dim humCfactor As Int16
        Dim barCalibrationRatio As Single
        Dim barCalibrationOffset As Single
    End Structure

    'TimeZoneSetting
    'Holds information about the time zone and daylight savings mode settings on the console

    'Holds information about the Time Zone and Daylight Savings settings on the VantagePro console. 
    'This structure is used by the functions GetTimeZoneSettings_V and SetTimeZoneSettings_V. 

    'The time zone can either be specified from a list (with bUseTimeZoneList = 1 and timeZone = a 
    ' constant from the Time Zone List) or by a GMT/UTC offset (with bUseTimeZoneList = 0 and 
    ' GMToffset = number of hours that clocks should be set ahead or behind GMT/UTC). The GMTofset 
    ' value is rounded to the nearest 15 minute (0.25 hour) time. The time zone set should be based 
    ' on the "Standard" time zone. Adjustments for daylight savings are taken into consideration by 
    ' the two daylight savings settings. 

    'Set bAutoDaylightSavings = 1 to have the console automatically switch Daylight Savings on or off. 
    ' The exact dates that this occurs vary depending on the latitude/longitude settings. See the 
    ' WeatherLink on-line help file for more details.

    'Set bAutoDaylightSavings = 0 to control daylight savings mode manually. (i.e. your location does 
    ' not observe daylight savings time, or does not use the same starting and ending dates that the 
    ' VantagePro console uses.) 

    'bDaylightSavingsOnNow will indicate whether daylight savings time is currently in effect.
    'If automatic daylight savings mode is selected, GetTimeZoneSettings_V will return the current status, 
    ' and the value is ignored by SetTimeZoneSettings_V.
    'If manual daylight savings mode is selected, the value is used by both GetTimeZoneSettings_V and 
    ' SetTimeZoneSettings_V. 

    'Note: Current versions of the VantagePro firmware (as of October 2005) do not account for the recent 
    ' changes in the daylight savings starting and ending dates in the US that are scheduled to take effect 
    ' fall 2007. This change should not effect European and Australian weather stations. 
    Structure TimeZoneSetting
        Dim bUseTimeZoneList As Int16
        Dim timeZone As Int16
        Dim GMToffset As Single
        Dim bAutoDaylightSavings As Int16
        Dim bDaylightSavingsOnNow As Int16
    End Structure

    'Current Active Alarm Fields

    'Retrieves the current bit field alarm status. 
    Class ActiveAlarmFields
        Dim insideAlarms As Byte
        Dim rainAlarms As Byte
        Dim outsideAlarms As UInt16
        Dim tempHumAlarms(8) As Byte
        Dim leafSoilAlarms(4) As Byte
    End Class

    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetCommTimeoutVal_V(ByVal readTimeout As Int16, ByVal writeTimeout As Int16) As UInt16
    End Function


    'Initialization Functions
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetDllVersion_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function OpenCommPort_V(ByVal comPort As Int16, ByVal baudRate As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function OpenDefaultCommPort() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function CloseCommPort_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
        Public Function OpenUSBPort_V(ByVal usbSerialNumber As UInt32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
        Public Function CloseUSBPort_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetUSBDevSerialNumber_V() As UInt32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function OpenTCPIPPort_V(ByVal macAddress As String, ByVal ipAddress As String) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function CloseTCPIPPort_V() As Int16
    End Function


    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetUnits(ByVal Units As WeatherUnits) As Integer
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetUnits(ByVal Units As WeatherUnits) As Integer
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetRainCollectorModel(ByVal rainCModel As Byte) As Integer
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetRainCollectorModel() As Byte
    End Function



    'Lowlevel Functions
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetSerialChar() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutSerialStr(ByVal stringtoport As Byte) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutSerialChar(ByVal c As Byte) As Int16
    End Function


    'Station Configuration Functions

    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function InitStation_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetStationTime_V(ByRef dateTimeStamp As DateTimeStamp) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetStationTime_V(ByRef dateTimeStamp As DateTimeStamp) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetArchivePeriod_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetArchivePeriod_V(ByVal intervalCode As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutTotalRain_V(ByVal TotalRain As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetStationFirmwareDate_V(ByRef timeStamp As DateTimeStamp) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetReceptionData_V(ByRef receptionStats As ReceptionStats) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetVantageLat_V(ByRef latitude As LatLonValue) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetVantageLat_V(ByRef latitude As LatLonValue) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetVantageLon_V(ByRef longtitude As LatLonValue) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetVantageLon_V(ByRef longtitude As LatLonValue) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetVantageLamp_V(ByVal ampState As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetNewBaud_V(ByVal baud As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetVantageTxConfig_V(ByRef txConfig As TxConfiguration) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetVantageTxConfig_V(ByRef txConfig As TxConfiguration) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetBarometerData_V(ByRef barCalData As BarCalData) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetRainCollectorModelOnStation_V(ByVal rainCModel As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetRainCollectorModelOnStation_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetAndSetRainCollectorModelOnStation_V() As Int16
    End Function



    'Current Data Functions
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function LoadCurrentVantageData_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetBarometer() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetOutsideTemp_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetDewPt() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetWindChill() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetInsideTemp_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetInsideHumidity_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetOutsideHumidity_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetTotalRain_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetDailyRain_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetStormRain_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetMonthlyRain_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetWindSpeed_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetWindDir_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetWindDirStr_V(ByVal dirStr As String) As String
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetRainRate_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetET_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetMonthlyET_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetYearlyET_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetSolarRad_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetUV_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHeatIndex_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetActiveAlarms_V(ByRef alarmFieldStruct As ActiveAlarmFields) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetCurrentDataByID_V(ByVal id As UInt32) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetCurrentDataStrByID_V(ByVal id As UInt32, ByVal s As String, ByVal bufferLength As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetStartOfCurrentStorm_V(ByRef dt As DateTimeStamp) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetSunriseTime_V(ByRef dt As DateTimeStamp) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetSunsetTime_V(ByRef dt As DateTimeStamp) As Int16
    End Function



    ' No such secsion -- move them
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetCommTimeoutVal_V(ByVal ReadTimeout As Integer, ByVal WriteTimeout As Integer) As Integer
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutBarometer_V(ByVal bar As Integer, ByVal elev As Integer) As Long
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetStationTime_V(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer, ByVal hour As Integer, ByVal min As Integer) As Integer
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetStationTime_V() As Integer
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutTotalRain_V(ByVal TotalRain As Single) As Int32
    End Function



    ' Alarm Functions
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function LoadVantageAlarms_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetVantageAlarms_V() As Int16
    End Function

    ' Get alarm values
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetBarRiseAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetBarFallAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetTimeAlarm_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetTimeAlarmStr_V() As String
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetInsideLowTempAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetInsideHiTempAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetOutsideLowTempAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetOutsideHiTempAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowInsideHumAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiInsideHumAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowOutsideHumAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiOutsideHumAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowWindChillAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowDewPtAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiDewPtAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiSolarRadAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiWindSpeedAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHi10MinWindSpeedAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiHeatIndexAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiTHSWAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiRainRateAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiDailyRainAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiRainStormAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetFlashFloodAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiUVAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiUVMedAlarm_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowExtraTempAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiExtraTempAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowExtraHumAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiExtraHumAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowSoilTempAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiSoilTempAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowSoilMoistureAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiSoilMoistureAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowLeafTempAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLeafTempAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowLeafWetAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLeafWetAlarm_V(ByVal sensorNumber As Int16) As Single
    End Function


    ' Put Alarm thresholds
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutBarRiseAlarm_V(ByVal barRiseAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutBarFallAlarm_V(ByVal barFallAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutTimeAlarm_V(ByVal timeAlarm As String) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutInsideLowTempAlarm_V(ByVal lowtempAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutInsideHiTempAlarm_V(ByVal hitempAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideLowTempAlarm_V(ByVal lowtempAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideHiTempAlarm_V(ByVal hitempAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowInsideHumAlarm_V(ByVal lowInsideAlarm As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiInsideHumAlarm_V(ByVal hiInsideAlarm As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowOutsideHumAlarm_V(ByVal lowOutsideAlarm As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiOutsideHumAlarm_V(ByVal hiOutsideAlarm As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowWindChillAlarm_V(ByVal lowWindChillAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowDewPtAlarm_V(ByVal lowDewPoint As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiDewPtAlarm_V(ByVal hiDewPoint As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiSolarRadAlarm_V(ByVal solarAlarm As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiWindSpeedAlarm_V(ByVal hiAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHi10MinWindSpeedAlarm_V(ByVal hiAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiHeatIndexAlarm_V(ByVal heatAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiTHSWAlarm_V(ByVal thswAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiRainFloodAlarm_V(ByVal hiAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutRainPerDayAlarm_V(ByVal hiAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutRainStormAlarm_V(ByVal hiAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutRainRateAlarm_V(ByVal hiAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiUVAlarm_V(ByVal uvAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiUVMedAlarm_V(ByVal uvMedAlarm As Single) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearBarRiseAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearBarFallAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearTimeAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearInsideLowTempAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearInsideHiTempAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearOutsideLowTempAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearOutsideHiTempAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowInsideHumAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiInsideHumAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowOutsideHumAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiOutsideHumAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowWindChillAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowDewPtAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiDewPtAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiSolarRadAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiWindSpeedAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHi10MinWindSpeedAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiHeatIndexAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiTHSWAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiRainFloodAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiRainPerDayAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearRainStormAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearRainRateAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiUVAlarm_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiUVMedAlarm_V() As Int16
    End Function


    ' HiLow Functions
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function LoadVantageHiLows_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiOutsideTemp_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowOutsideTemp_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowTimesOutTemp_V(ByRef DateTimeHiOutTemp As DateTime, ByRef DateTimeLowOutTemp As DateTime) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiInsideTemp_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowInsideTemp_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowTimesInTemp_V(ByRef DateTimeHiInTemp As DateTime, ByRef DateTimeLowInTemp As DateTime) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiOutsideHum_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowOutsideHum_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowTimesOutHum_V(ByRef DateTimeHiOutHum As DateTime, ByRef DateTimeLowOutHum As DateTime) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiInsideHum_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowInsideHum_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowTimesInHum_V(ByRef DateTimeHiInHum As DateTime, ByRef DateTimeLowInHum As DateTime) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiDewPt_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowDewPt_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowTimesDewPt_V(ByRef DateTimeHiDewPt As DateTime, ByVal DateTimeLowDewPt As DateTime) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowWindChill_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetLowTimesWindChill_V(ByVal DateTimeLowWindChill As DateTime) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiWindSpeed_V() As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiTimesWindSpeed_V(ByVal DateTimeHiWindSpeed As DateTime) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowDataByID_V(ByVal weatherDataID As UInt32) As Single
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowDataStrByID_V(ByVal weatherDataID As UInt32, ByVal s As String, ByVal bufferLength As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowTimeByID_V(ByVal weatherDataID As UInt32, ByRef dateTimeValue As DateTimeStamp) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetHiLowTimeStrByID_V(ByVal weatherDataID As UInt32, ByVal s As String, ByVal bufferLength As Int16) As Int16
    End Function

    ' Calibrate Functions
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function LoadVantageCalibration_V(ByRef vantageCalibration As CurrentVantageCalibration) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideTempCalibrationValue_V(ByVal calValue As Single) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutInsideTempCalibrationValue_V(ByVal calValue As Single) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideHumCalibrationValue_V(ByVal calValue As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutInsideHumCalibrationValue_V(ByVal calValue As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideTempCalibrationOffset_V(ByVal calValue As Single) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutInsideTempCalibrationOffset_V(ByVal calValue As Single) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideHumCalibrationOffset_V(ByVal calValue As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutInsideHumCalibrationOffset_V(ByVal calValue As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideTempCalibrationValueEx_V(ByVal sensorNumber As Int32, ByVal calValue As Single) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideTempCalibrationOffsetEx_V(ByVal sensorNumber As Int32, ByVal calValue As Single) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideHumCalibrationValueEx_V(ByVal sensorNumber As Int32, ByVal calValue As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutOutsideHumCalibrationOffsetEx_V(ByVal sensorNumber As Int32, ByVal calValue As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function SetVantageCalibration_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutBarometer_V(ByVal bar As Single, ByVal elev As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutWindDirCalibrationOffset_V(ByVal windDirCal As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetWindDirCalibrationOffset_V(ByRef windDirCal As Int16) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowExtraTempAlarm_V(ByVal sensorNumber As Int32, ByVal lowtempExAlarm As Single) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiExtraTempAlarm_V(ByVal sensorNumber As Int32, ByVal hitempExAlarm As Single) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowExtraHumAlarm_V(ByVal sensorNumber As Int32, ByVal lowHumExAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiExtraHumAlarm_V(ByVal sensorNumber As Int32, ByVal hiHumExAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowSoilTempAlarm_V(ByVal sensorNumber As Int32, ByVal lowSoilTempAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiSoilTempAlarm_V(ByVal sensorNumber As Int32, ByVal hiSoilTempAlarmAs As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowSoilMoistureAlarm_V(ByVal sensorNumber As Int32, ByVal lowSoilMoistureAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiSoilMoistureAlarm_V(ByVal sensorNumber As Int32, ByVal hiSoilMoistureAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowLeafTempAlarm_V(ByVal sensorNumber As Int32, ByVal lowLeafTempAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiLeafTempAlarm_V(ByVal sensorNumber As Int32, ByVal hiLeafTempAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutLowLeafWetAlarm_V(ByVal sensorNumber As Int32, ByVal lowLeafWetAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function PutHiLeafWetAlarm_V(ByVal sensorNumber As Int32, ByVal hiLeafWetAlarm As Int32) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowExtraTempAlarm_V(ByVal sensorNumber As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiExtraTempAlarm_V(ByVal sensorNumber As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowExtraHumAlarm_V(ByVal sensorNumber As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiExtraHumAlarm_V(ByVal sensorNumber As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowSoilTempAlarm_V(ByVal sensorNumber As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiSoilTempAlarm_V(ByVal sensorNumber As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowSoilMoistureAlarm_V(ByVal sensorNumber As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiSoilMoistureAlarm_V(ByVal sensorNumber As Int32) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowLeafTempAlarm_V(ByVal sensorNumber As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiLeafTempAlarm_V(ByVal sensorNumber As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearLowLeafWetAlarm_V(ByVal sensorNumber As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearHiLeafWetAlarm_V(ByVal sensorNumber As Int16) As Int16
    End Function


    ' Download functions
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function DownloadData_V(ByVal dateTimeStamp As DateTimeStamp) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function DownloadWebData_V(ByVal dateTimeStamp As DateTimeStamp, ByVal userName As String, ByVal password As String) As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetNumberOfArchiveRecords_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetMemoryArchiveRecordCount_V() As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetMemoryArchiveCountAfterDate_V(ByRef dateTimeStamp As DateTimeStamp) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetArchiveRecord_V(ByRef newRecord As WeatherRecordStruct, ByVal i As Int16) As Int16
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function GetArchiveRecordEx_V(ByRef newRecordStruct As WeatherRecordStructEx, ByVal i As Int16) As Int16
    End Function

    ' Clear functions
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearVantageLows_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearVantageAlarms_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearVantageCalNums_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearCurrentData_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearStoredData_V() As Int32
    End Function
    <DllImport("VantagePro.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.StdCall)> _
    Public Function ClearReceiveBuffer_V() As Int32
    End Function

End Module
