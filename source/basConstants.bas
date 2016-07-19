Attribute VB_Name = "basConstants"
'------------------------------------------------------------------------
' Description  : contains all global constants
'------------------------------------------------------------------------

' Copyright 2016 Matthias Carell
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

'Options
Option Explicit

'log level (range is 1 to 100)
Global Const cLogDebug = 100
Global Const cLogInfo = 90
Global Const cLogWarning = 50
Global Const cLogError = 30
Global Const cLogCritical = 1

'current log level - decreasing log level means decreasing amount of messages
Global Const cCurrentLogLevel = 90

'data table column types
Global Const cColTypeDomain = "domain"
Global Const cColTypeAggregate = "aggregate"
Global Const cColTypeFeature = "feature"
Global Const cColTypeScenario = "scenario"
