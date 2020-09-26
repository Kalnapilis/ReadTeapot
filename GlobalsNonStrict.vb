Option Strict Off
Imports System.IO
Imports System.Threading

Module GlobalsNonStrict ' Option Strict is Off
    Public Function CanReadFrom(ByVal DirectoryPath As String) As Boolean
        Dim dirInfo As System.IO.DirectoryInfo = FileIO.FileSystem.GetDirectoryInfo(DirectoryPath)
        Dim currentUser As System.Security.Principal.WindowsIdentity = _
           System.Security.Principal.WindowsIdentity.GetCurrent()
        Dim currentPrinciple As System.Security.Principal.WindowsPrincipal = _
           CType(System.Threading.Thread.CurrentPrincipal, System.Security.Principal.WindowsPrincipal)
        Dim acl As System.Security.AccessControl.AuthorizationRuleCollection = _
           dirInfo.GetAccessControl().GetAccessRules(True, True, GetType(System.Security.Principal.SecurityIdentifier))
        Dim currentRule As System.Security.AccessControl.FileSystemAccessRule
        Dim denyread As Boolean = False
        Dim allowread As Boolean = False
        For x As Integer = 0 To acl.Count - 1
            currentRule = acl(x)
            If currentUser.User.Equals(currentRule.IdentityReference) Or currentPrinciple.IsInRole(currentRule.IdentityReference) Then
                If currentRule.AccessControlType.Equals(System.Security.AccessControl.AccessControlType.Deny) Then
                    If (currentRule.FileSystemRights And System.Security.AccessControl.FileSystemRights.Read) = System.Security.AccessControl.FileSystemRights.Read Then denyread = True
                Else
                    If currentRule.AccessControlType.Equals(System.Security.AccessControl.AccessControlType.Allow) Then allowread = True
                End If
            End If
        Next
        If allowread And Not (denyread) Then
            Return True
        Else
            Return False
        End If
    End Function
   Public Function CanWriteTo(ByVal DirectoryPath As String) As Boolean
      Dim dirInfo As System.IO.DirectoryInfo = FileIO.FileSystem.GetDirectoryInfo(DirectoryPath)
      Dim currentUser As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent()
      Dim currentPrinciple As System.Security.Principal.WindowsPrincipal = System.Threading.Thread.CurrentPrincipal
      Dim acl As System.Security.AccessControl.AuthorizationRuleCollection = dirInfo.GetAccessControl().GetAccessRules(True, True, GetType(System.Security.Principal.SecurityIdentifier))
      Dim currentRule As System.Security.AccessControl.FileSystemAccessRule
      Dim denyread As Boolean = False
      Dim allowread As Boolean = False
      For x As Integer = 0 To acl.Count - 1
         currentRule = acl(x)
         If currentUser.User.Equals(currentRule.IdentityReference) Or currentPrinciple.IsInRole(currentRule.IdentityReference) Then
            If currentRule.AccessControlType.Equals(System.Security.AccessControl.AccessControlType.Deny) Then
               If (currentRule.FileSystemRights And System.Security.AccessControl.FileSystemRights.Write) = System.Security.AccessControl.FileSystemRights.Write Then denyread = True
            Else
               If currentRule.AccessControlType.Equals(System.Security.AccessControl.AccessControlType.Allow) Then allowread = True
            End If
         End If
      Next
      If allowread And Not (denyread) Then
         Return True
      Else
         Return False
      End If
   End Function
   Public Class SunData

      Private Shared Function DegMinToFractional(ByVal DDMM As Single) As Single
1850:    ' This subroutine converts DD.MM input to DD.DD
         Dim DEGTMP As Single
         DEGTMP = (System.Math.Abs(DDMM) - System.Math.Abs(Fix(DDMM))) * 100 / 60
         Return (Fix(System.Math.Abs(DDMM)) + DEGTMP) * System.Math.Sign(DDMM)
      End Function
      Private Shared Function ConvertFractionalTime(ByVal IsSunset As Boolean, _
                          ByVal ObsTimeIn As Single) As Date
         Dim sngObsHour, sngObsMinute As Single
1760:    sngObsHour = Int(ObsTimeIn)
         sngObsMinute = ObsTimeIn - sngObsHour
         sngObsMinute = Int((sngObsMinute * 600 + 5) / 10) 'get rid of seconds
         If IsSunset Then
            Return TimeSerial(CInt(sngObsHour) + 12, CInt(sngObsMinute), 0)
         Else
            Return TimeSerial(CInt(sngObsHour), CInt(sngObsMinute), 0)
         End If
      End Function

      Public Shared Sub Sun(ByRef SunRise As Date, ByRef SunSet As Date, ByVal ObsDate As Date, _
                            ByVal Latitude As Single, ByVal Longitude As Single)
         'Return sunrise and sunset
         ' One note regarding azimuth angles: In the SOUTHERN hemisphere,
         ' this program assumes that the South Pole is zero degrees, and
         ' the azimuth angle is measured COUNTER-clockwise through East.
         ' Therefore, an azimuth angle of 108 degrees is NORTH of East.
         ' When Lat and Long are input, use a negative value for Southern
         ' Latitudes and for Eastern Longitudes.
         ' ----------------- Program Begins ---------------------------
         '"This program finds the declination of the sun, the equation"
         '"of time, the azimuth angles of sunrise and sunset, and the"
         '"times of sunrise and sunset for any point on earth."
         ' Output times are in the local time zone with Daylight Saving adjustment
         '"Input eastern longitudes and southern latitudes as NEGATIVE."
         ' INPUT"ENTER LATITUDE (FORMAT DD.MM)";D1
         ' INPUT"ENTER LONGITUDE (FORMAT DD.MM)";D2

         Dim RadiansPerWeek As Single = 3.1415926536 / 26
         Dim DegPerRadian As Single = 57.29577951
         Dim LongitudeDecimal As Single
         Dim LongitudeTZDelta As Single
         Dim WeekOfYear As Single
         Dim DeclinationOfSun As Single
         Dim EquationOfTime As Single
         Dim LatitudeCosine As Single
         Dim DeclinationCoSine As Single
         Dim DeclinationSine As Single
         Dim Y As Single
         Dim AzimuthOfSunrise As Single
         Dim st As Single
         Dim RawTime As Single
         Dim TT As Single
         Dim CT As Single
         Dim LatitudeDecimal, d1ddmm, d2ddmm As Single
         Dim T3 As Single

         d1ddmm = Latitude : d2ddmm = Longitude
         If Latitude < 0 Then Latitude = Latitude + 180
         If d2ddmm < 0 Then d2ddmm = d2ddmm + 360
         LatitudeDecimal = DegMinToFractional(d1ddmm)
         LongitudeDecimal = DegMinToFractional(d2ddmm)
         T3 = Fix(LongitudeDecimal / 15) * 15 ' finds time zone beginning
         LongitudeTZDelta = (LongitudeDecimal - T3) / 15
         WeekOfYear = ObsDate.DayOfYear / 7

         DeclinationOfSun = 0.4560001 - 22.195 * System.Math.Cos(RadiansPerWeek * WeekOfYear) _
                        - 0.43 * System.Math.Cos(2 * RadiansPerWeek * WeekOfYear) _
                        - 0.156 * System.Math.Cos(3 * RadiansPerWeek * WeekOfYear) _
                        + 3.83 * System.Math.Sin(RadiansPerWeek * WeekOfYear) _
                        + 0.06 * System.Math.Sin(2 * RadiansPerWeek * WeekOfYear) _
                        - 0.082 * System.Math.Sin(3 * RadiansPerWeek * WeekOfYear)

         EquationOfTime = 0.008000001 _
                        + 0.51 * System.Math.Cos(RadiansPerWeek * WeekOfYear) _
                        - 3.197 * System.Math.Cos(2 * RadiansPerWeek * WeekOfYear) _
                        - 0.106 * System.Math.Cos(3 * RadiansPerWeek * WeekOfYear) _
                        - 0.15 * System.Math.Cos(4 * RadiansPerWeek * WeekOfYear) _
                        - 7.317001 * System.Math.Sin(RadiansPerWeek * WeekOfYear) _
                        - 9.471001 * System.Math.Sin(2 * RadiansPerWeek * WeekOfYear) _
                        - 0.391 * System.Math.Sin(3 * RadiansPerWeek * WeekOfYear) _
                        - 0.242 * System.Math.Sin(4 * RadiansPerWeek * WeekOfYear)
         LatitudeCosine = System.Math.Cos(Latitude / DegPerRadian)
         DeclinationSine = System.Math.Sin(DeclinationOfSun / DegPerRadian)
         DeclinationCoSine = System.Math.Cos(DeclinationOfSun / DegPerRadian)
         Y = DeclinationSine / LatitudeCosine
         If System.Math.Abs(Y) >= 1 Then
            End ' #### throw an exception
         End If '"NO SUNRISE OR SUNSET"
         AzimuthOfSunrise = 90 - DegPerRadian * System.Math.Atan(Y / System.Math.Sqrt(1 - Y * Y))
         '"AZIMUTH OF SUNSET: 360-ABS(AzimuthOfSunrise);
         st = System.Math.Sin(AzimuthOfSunrise / DegPerRadian) / DeclinationCoSine
         If System.Math.Abs(st) >= 1 Then
            RawTime = 6
            TT = 6
         Else
            CT = System.Math.Sqrt(1 - st * st)
            RawTime = DegPerRadian / 15 * System.Math.Atan(st / CT)
            TT = RawTime
         End If
1660:    If DeclinationOfSun < 0 And Latitude < 90 Then RawTime = 12 - RawTime : TT = RawTime
         If DeclinationOfSun > 0 And Latitude > 90 Then RawTime = 12 - RawTime : TT = RawTime
         RawTime = RawTime + LongitudeTZDelta - EquationOfTime / 60 - 0.04

         SunRise = ConvertFractionalTime(False, RawTime)

         RawTime = 12 - TT : RawTime = RawTime + LongitudeTZDelta - EquationOfTime / 60 + 0.04
         SunSet = ConvertFractionalTime(True, RawTime)

         Dim LocalZone As TimeZone = TimeZone.CurrentTimeZone
         If TimeZone.IsDaylightSavingTime(ObsDate, LocalZone.GetDaylightChanges(ObsDate.Year)) Then
            SunRise = DateAdd(DateInterval.Hour, 1, SunRise)
            SunSet = DateAdd(DateInterval.Hour, 1, SunSet)
         End If
      End Sub

   End Class ' Sun Data
End Module
