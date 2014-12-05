# -*- coding: mbcs -*-
typelib_path = 'mstask.dll'
_lcid = 0 # change this if required
from ctypes import *
from comtypes import GUID
from comtypes import IUnknown
from ctypes import HRESULT
from comtypes import helpstring
from comtypes import COMMETHOD
from comtypes import dispid
WSTRING = c_wchar_p
from comtypes import wireHWND
from comtypes import CoClass
from comtypes import GUID
from comtypes import IUnknown

################################################################
### Constants
################################################################

# These should probably be coming from elsewhere!
SCHED_S_TASK_READY                  = 0x00041300 # The task is ready to run at its next scheduled time.
SCHED_S_TASK_RUNNING                = 0x00041301 # The task is currently running.
SCHED_S_TASK_DISABLED               = 0x00041302 # The task will not run at the scheduled times because it has been disabled.
SCHED_S_TASK_HAS_NOT_RUN            = 0x00041303 # The task has not yet run.
SCHED_S_TASK_NO_MORE_RUNS           = 0x00041304 # There are no more runs scheduled for this task.
SCHED_S_TASK_NOT_SCHEDULED          = 0x00041305 # One or more of the properties that are needed to run this task on a schedule have not been set.
SCHED_S_TASK_TERMINATED             = 0x00041306 # The last run of the task was terminated by the user.
SCHED_S_TASK_NO_VALID_TRIGGERS      = 0x00041307 # Either the task has no triggers or the existing triggers are disabled or not set.
SCHED_S_EVENT_TRIGGER               = 0x00041308 # Event triggers do not have set run times.
SCHED_E_TRIGGER_NOT_FOUND           = 0x80041309 # A task's trigger is not found.
SCHED_E_TASK_NOT_READY              = 0x8004130A # One or more of the properties required to run this task have not been set.
SCHED_E_TASK_NOT_RUNNING            = 0x8004130B # There is no running instance of the task.
SCHED_E_SERVICE_NOT_INSTALLED       = 0x8004130C # The Task Scheduler service is not installed on this computer.
SCHED_E_CANNOT_OPEN_TASK            = 0x8004130D # The task object could not be opened.
SCHED_E_INVALID_TASK                = 0x8004130E # The object is either an invalid task object or is not a task object.
SCHED_E_ACCOUNT_INFORMATION_NOT_SET = 0x8004130F # No account information could be found in the Task Scheduler security database for the task indicated.
SCHED_E_ACCOUNT_NAME_NOT_FOUND      = 0x80041310 # Unable to establish existence of the account specified.
SCHED_E_ACCOUNT_DBASE_CORRUPT       = 0x80041311 # Corruption was detected in the Task Scheduler security database; the database has been reset.
SCHED_E_NO_SECURITY_SERVICES        = 0x80041312 # Task Scheduler security services are available only on Windows NT.
SCHED_E_UNKNOWN_OBJECT_VERSION      = 0x80041313 # The task object version is either unsupported or invalid.
SCHED_E_UNSUPPORTED_ACCOUNT_OPTION  = 0x80041314 # The task has been configured with an unsupported combination of account settings and run time options.
SCHED_E_SERVICE_NOT_RUNNING         = 0x80041315 # The Task Scheduler Service is not running.
SCHED_E_UNEXPECTEDNODE              = 0x80041316 # The task XML contains an unexpected node.
SCHED_E_NAMESPACE                   = 0x80041317 # The task XML contains an element or attribute from an unexpected namespace.
SCHED_E_INVALIDVALUE                = 0x80041318 # The task XML contains a value which is incorrectly formatted or out of range.
SCHED_E_MISSINGNODE                 = 0x80041319 # The task XML is missing a required element or attribute.
SCHED_E_MALFORMEDXML                = 0x8004131A # The task XML is malformed.
SCHED_S_SOME_TRIGGERS_FAILED        = 0x0004131B # The task is registered, but not all specified triggers will start the task.
SCHED_S_BATCH_LOGON_PROBLEM         = 0x0004131C # The task is registered, but may fail to start. Batch logon privilege needs to be enabled for the task principal.
SCHED_E_TOO_MANY_NODES              = 0x8004131D # The task XML contains too many nodes of the same type.
SCHED_E_PAST_END_BOUNDARY           = 0x8004131E # The task cannot be started after the trigger end boundary.
SCHED_E_ALREADY_RUNNING             = 0x8004131F # An instance of this task is already running.
SCHED_E_USER_NOT_LOGGED_ON          = 0x80041320 # The task will not run because the user is not logged on.
SCHED_E_INVALID_TASK_HASH           = 0x80041321 # The task image is corrupt or has been tampered with.
SCHED_E_SERVICE_NOT_AVAILABLE       = 0x80041322 # The Task Scheduler service is not available.
SCHED_E_SERVICE_TOO_BUSY            = 0x80041323 # The Task Scheduler service is too busy to handle your request. Please try again later.
SCHED_E_TASK_ATTEMPTED              = 0x80041324 # The Task Scheduler service attempted to run the task, but the task did not run due to one of the constraints in the task definition.
SCHED_S_TASK_QUEUED                 = 0x00041325 # The Task Scheduler service has asked the task to run.
SCHED_E_TASK_DISABLED               = 0x80041326 # The task is disabled.
SCHED_E_TASK_NOT_V1_COMPAT          = 0x80041327 # The task has properties that are not compatible with earlier versions of Windows.
SCHED_E_START_ON_DEMAND             = 0x80041328 # The task settings do not allow the task to start on demand.
NORMAL_PRIORITY_CLASS	=  32
IDLE_PRIORITY_CLASS	=  64
HIGH_PRIORITY_CLASS	= 128
REALTIME_PRIORITY_CLASS	= 256

TASK_FIRST_WEEK = 1
TASK_SECOND_WEEK = 2
TASK_THIRD_WEEK = 3
TASK_FOURTH_WEEK = 4
TASK_LAST_WEEK = 5

TASK_SUNDAY = 1
TASK_MONDAY = 2
TASK_TUESDAY = 4
TASK_WEDNESDAY = 8
TASK_THURSDAY = 16
TASK_FRIDAY = 32
TASK_SATURDAY = 64

TASK_FLAG_INTERACTIVE = 1
TASK_FLAG_DELETE_WHEN_DONE = 2
TASK_FLAG_DISABLED = 4
TASK_FLAG_START_ONLY_IF_IDLE = 16
TASK_FLAG_KILL_ON_IDLE_END = 32
TASK_FLAG_DONT_START_IF_ON_BATTERIES = 64
TASK_FLAG_KILL_IF_GOING_ON_BATTERIES = 128
TASK_FLAG_RUN_ONLY_IF_DOCKED = 256
TASK_FLAG_HIDDEN = 512
TASK_FLAG_RUN_IF_CONNECTED_TO_INTERNET = 1024
TASK_FLAG_RESTART_ON_IDLE_RESUME = 2048
TASK_FLAG_SYSTEM_REQUIRED = 4096
TASK_FLAG_RUN_ONLY_IF_LOGGED_ON = 8192

TASK_JANUARY = 1
TASK_FEBRUARY = 2
TASK_MARCH = 4
TASK_APRIL = 8
TASK_MAY = 16
TASK_JUNE = 32
TASK_JULY = 64
TASK_AUGUST = 128
TASK_SEPTEMBER = 256
TASK_OCTOBER = 512
TASK_NOVEMBER = 1024
TASK_DECEMBER = 2048

TASK_TRIGGER_FLAG_HAS_END_DATE = 1
TASK_TRIGGER_FLAG_KILL_AT_DURATION_END = 2
TASK_TRIGGER_FLAG_DISABLED = 4

# 1440 = 60 mins/hour * 24 hrs/day since a trigger/TASK could run all day at
# one minute intervals.
TASK_MAX_RUN_TIMES = 1440

################################################################
### Structures
################################################################

# This should probably be coming from elsewhere!
class _SYSTEMTIME(Structure):
    _fields_ = [
        ('wYear', c_ushort),
        ('wMonth', c_ushort),
        ('wDayOfWeek', c_ushort),
        ('wDay', c_ushort),
        ('wHour', c_ushort),
        ('wMinute', c_ushort),
        ('wSecond', c_ushort),
        ('wMilliseconds', c_ushort),
    ]
assert sizeof(_SYSTEMTIME) == 16, sizeof(_SYSTEMTIME)
assert alignment(_SYSTEMTIME) == 2, alignment(_SYSTEMTIME)

class _DAILY(Structure):
    _fields_ = [
        ('DaysInterval', c_ushort),
    ]
assert sizeof(_DAILY) == 2, sizeof(_DAILY)
assert alignment(_DAILY) == 2, alignment(_DAILY)

class _WEEKLY(Structure):
    _fields_ = [
        ('WeeksInterval', c_ushort),
        ('rgfDaysOfTheWeek', c_ushort),
    ]
assert sizeof(_WEEKLY) == 4, sizeof(_WEEKLY)
assert alignment(_WEEKLY) == 2, alignment(_WEEKLY)

class _MONTHLYDATE(Structure):
    _fields_ = [
        ('rgfDays', c_ulong),
        ('rgfMonths', c_ushort),
    ]
assert sizeof(_MONTHLYDATE) == 8, sizeof(_MONTHLYDATE)
assert alignment(_MONTHLYDATE) == 4, alignment(_MONTHLYDATE)

class _MONTHLYDOW(Structure):
    _fields_ = [
        ('wWhichWeek', c_ushort),
        ('rgfDaysOfTheWeek', c_ushort),
        ('rgfMonths', c_ushort),
    ]
assert sizeof(_MONTHLYDOW) == 6, sizeof(_MONTHLYDOW)
assert alignment(_MONTHLYDOW) == 2, alignment(_MONTHLYDOW)

class _TRIGGER_TYPE_UNION(Union):
    _fields_ = [
        ('Daily', _DAILY),
        ('Weekly', _WEEKLY),
        ('MonthlyDate', _MONTHLYDATE),
        ('MonthlyDOW', _MONTHLYDOW),
    ]
assert sizeof(_TRIGGER_TYPE_UNION) == 8, sizeof(_TRIGGER_TYPE_UNION)
assert alignment(_TRIGGER_TYPE_UNION) == 4, alignment(_TRIGGER_TYPE_UNION)

# values for enumeration '_TASK_TRIGGER_TYPE'
TASK_TIME_TRIGGER_ONCE = 0
TASK_TIME_TRIGGER_DAILY = 1
TASK_TIME_TRIGGER_WEEKLY = 2
TASK_TIME_TRIGGER_MONTHLYDATE = 3
TASK_TIME_TRIGGER_MONTHLYDOW = 4
TASK_EVENT_TRIGGER_ON_IDLE = 5
TASK_EVENT_TRIGGER_AT_SYSTEMSTART = 6
TASK_EVENT_TRIGGER_AT_LOGON = 7
_TASK_TRIGGER_TYPE = c_int # enum

class _TASK_TRIGGER(Structure):
    _fields_ = [
        ('cbTriggerSize', c_ushort),
        ('Reserved1', c_ushort),
        ('wBeginYear', c_ushort),
        ('wBeginMonth', c_ushort),
        ('wBeginDay', c_ushort),
        ('wEndYear', c_ushort),
        ('wEndMonth', c_ushort),
        ('wEndDay', c_ushort),
        ('wStartHour', c_ushort),
        ('wStartMinute', c_ushort),
        ('MinutesDuration', c_ulong),
        ('MinutesInterval', c_ulong),
        ('rgFlags', c_ulong),
        ('TriggerType', _TASK_TRIGGER_TYPE),
        ('Type', _TRIGGER_TYPE_UNION),
        ('Reserved2', c_ushort),
        ('wRandomMinutesInterval', c_ushort),
    ]
assert sizeof(_TASK_TRIGGER) == 48, sizeof(_TASK_TRIGGER)
assert alignment(_TASK_TRIGGER) == 4, alignment(_TASK_TRIGGER)

################################################################
### COM Interfaces
################################################################

################################################################
##class ITaskTrigger(object):
##    def GetTrigger(self):
##        '-no docstring-'
##        #return pTrigger
##
##    def GetTriggerString(self):
##        '-no docstring-'
##        #return ppwszTrigger
##
##    def SetTrigger(self, pTrigger):
##        '-no docstring-'
##        #return 
##
class ITaskTrigger(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{148BD52B-A2AB-11CE-B11F-00AA00530503}')
    _idlflags_ = []
    _methods_ = [
        COMMETHOD([], HRESULT, 'SetTrigger',
                  ( ['in'], POINTER(_TASK_TRIGGER), 'pTrigger' )),
        COMMETHOD([], HRESULT, 'GetTrigger',
                  ( ['out'], POINTER(_TASK_TRIGGER), 'pTrigger' )),
        COMMETHOD([], HRESULT, 'GetTriggerString',
                  ( ['out'], POINTER(WSTRING), 'ppwszTrigger' )),
    ]

################################################################
##class IScheduledWorkItem(object):
##    def SetWorkItemData(self, cbData, rgbData):
##        '-no docstring-'
##        #return 
##
##    def GetWorkItemData(self):
##        '-no docstring-'
##        #return pcbData, prgbData
##
##    def SetCreator(self, pwszCreator):
##        '-no docstring-'
##        #return 
##
##    def SetComment(self, pwszComment):
##        '-no docstring-'
##        #return 
##
##    def GetExitCode(self):
##        '-no docstring-'
##        #return pdwExitCode
##
##    def EditWorkItem(self, hParent, dwReserved):
##        '-no docstring-'
##        #return 
##
##    def CreateTrigger(self):
##        '-no docstring-'
##        #return piNewTrigger, ppTrigger
##
##    def GetRunTimes(self, pstBegin, pstEnd):
##        '-no docstring-'
##        #return pCount, rgstTaskTimes
##
##    def DeleteTrigger(self, iTrigger):
##        '-no docstring-'
##        #return 
##
##    def GetTriggerCount(self):
##        '-no docstring-'
##        #return pwCount
##
##    def GetIdleWait(self):
##        '-no docstring-'
##        #return pwIdleMinutes, pwDeadlineMinutes
##
##    def GetTriggerString(self, iTrigger):
##        '-no docstring-'
##        #return ppwszTrigger
##
##    def Run(self):
##        '-no docstring-'
##        #return 
##
##    def GetFlags(self):
##        '-no docstring-'
##        #return pdwFlags
##
##    def SetErrorRetryInterval(self, wRetryInterval):
##        '-no docstring-'
##        #return 
##
##    def GetCreator(self):
##        '-no docstring-'
##        #return ppwszCreator
##
##    def SetIdleWait(self, wIdleMinutes, wDeadlineMinutes):
##        '-no docstring-'
##        #return 
##
##    def GetAccountInformation(self):
##        '-no docstring-'
##        #return ppwszAccountName
##
##    def GetStatus(self):
##        '-no docstring-'
##        #return phrStatus
##
##    def SetErrorRetryCount(self, wRetryCount):
##        '-no docstring-'
##        #return 
##
##    def SetFlags(self, dwFlags):
##        '-no docstring-'
##        #return 
##
##    def GetErrorRetryInterval(self):
##        '-no docstring-'
##        #return pwRetryInterval
##
##    def Terminate(self):
##        '-no docstring-'
##        #return 
##
##    def SetAccountInformation(self, pwszAccountName, pwszPassword):
##        '-no docstring-'
##        #return 
##
##    def GetNextRunTime(self):
##        '-no docstring-'
##        #return pstNextRun
##
##    def GetMostRecentRunTime(self):
##        '-no docstring-'
##        #return pstLastRun
##
##    def GetComment(self):
##        '-no docstring-'
##        #return ppwszComment
##
##    def GetTrigger(self, iTrigger):
##        '-no docstring-'
##        #return ppTrigger
##
##    def GetErrorRetryCount(self):
##        '-no docstring-'
##        #return pwRetryCount
##
class IScheduledWorkItem(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{A6B952F0-A4B1-11D0-997D-00AA006887EC}')
    _idlflags_ = []
    _methods_ = [
        COMMETHOD([], HRESULT, 'CreateTrigger',
                  ( ['out'], POINTER(c_ushort), 'piNewTrigger' ),
                  ( ['out'], POINTER(POINTER(ITaskTrigger)), 'Trigger' )),
        COMMETHOD([], HRESULT, 'DeleteTrigger',
                  ( ['in'], c_ushort, 'iTrigger' )),
        COMMETHOD([], HRESULT, 'GetTriggerCount',
                  ( ['out'], POINTER(c_ushort), 'pwCount' )),
        COMMETHOD([], HRESULT, 'GetTrigger',
                  ( ['in'], c_ushort, 'iTrigger' ),
                  ( ['out'], POINTER(POINTER(ITaskTrigger)), 'Trigger' )),
        COMMETHOD([], HRESULT, 'GetTriggerString',
                  ( ['in'], c_ushort, 'iTrigger' ),
                  ( ['out'], POINTER(WSTRING), 'ppwszTrigger' )),
        COMMETHOD([], HRESULT, 'GetRunTimes',
                  ( ['in'], POINTER(_SYSTEMTIME), 'pstBegin' ),
                  ( ['in'], POINTER(_SYSTEMTIME), 'pstEnd' ),
                  ( ['in', 'out'], POINTER(c_ushort), 'pCount' ),
                  ( ['out'], POINTER(POINTER(_SYSTEMTIME)), 'rgstTaskTimes' )),
        COMMETHOD([], HRESULT, 'GetNextRunTime',
                  ( ['in', 'out'], POINTER(_SYSTEMTIME), 'pstNextRun' )),
        COMMETHOD([], HRESULT, 'SetIdleWait',
                  ( ['in'], c_ushort, 'wIdleMinutes' ),
                  ( ['in'], c_ushort, 'wDeadlineMinutes' )),
        COMMETHOD([], HRESULT, 'GetIdleWait',
                  ( ['out'], POINTER(c_ushort), 'pwIdleMinutes' ),
                  ( ['out'], POINTER(c_ushort), 'pwDeadlineMinutes' )),
        COMMETHOD([], HRESULT, 'Run'),
        COMMETHOD([], HRESULT, 'Terminate'),
        COMMETHOD([], HRESULT, 'EditWorkItem',
                  ( ['in'], wireHWND, 'hParent' ),
                  ( ['in'], c_ulong, 'dwReserved' )),
        COMMETHOD([], HRESULT, 'GetMostRecentRunTime',
                  ( ['out'], POINTER(_SYSTEMTIME), 'pstLastRun' )),
        COMMETHOD([], HRESULT, 'GetStatus',
                  ( ['out'], POINTER(HRESULT), 'phrStatus' )),
        COMMETHOD([], HRESULT, 'GetExitCode',
                  ( ['out'], POINTER(c_ulong), 'pdwExitCode' )),
        COMMETHOD([], HRESULT, 'SetComment',
                  ( ['in'], WSTRING, 'pwszComment' )),
        COMMETHOD([], HRESULT, 'GetComment',
                  ( ['out'], POINTER(WSTRING), 'ppwszComment' )),
        COMMETHOD([], HRESULT, 'SetCreator',
                  ( ['in'], WSTRING, 'pwszCreator' )),
        COMMETHOD([], HRESULT, 'GetCreator',
                  ( ['out'], POINTER(WSTRING), 'ppwszCreator' )),
        COMMETHOD([], HRESULT, 'SetWorkItemData',
                  ( ['in'], c_ushort, 'cbData' ),
                  ( ['in'], POINTER(c_ubyte), 'rgbData' )),
        COMMETHOD([], HRESULT, 'GetWorkItemData',
                  ( ['out'], POINTER(c_ushort), 'pcbData' ),
                  ( ['out'], POINTER(POINTER(c_ubyte)), 'prgbData' )),
        COMMETHOD([], HRESULT, 'SetErrorRetryCount',
                  ( ['in'], c_ushort, 'wRetryCount' )),
        COMMETHOD([], HRESULT, 'GetErrorRetryCount',
                  ( ['out'], POINTER(c_ushort), 'pwRetryCount' )),
        COMMETHOD([], HRESULT, 'SetErrorRetryInterval',
                  ( ['in'], c_ushort, 'wRetryInterval' )),
        COMMETHOD([], HRESULT, 'GetErrorRetryInterval',
                  ( ['out'], POINTER(c_ushort), 'pwRetryInterval' )),
        COMMETHOD([], HRESULT, 'SetFlags',
                  ( ['in'], c_ulong, 'dwFlags' )),
        COMMETHOD([], HRESULT, 'GetFlags',
                  ( ['out'], POINTER(c_ulong), 'pdwFlags' )),
        COMMETHOD([], HRESULT, 'SetAccountInformation',
                  ( ['in'], WSTRING, 'pwszAccountName' ),
                  ( ['in'], WSTRING, 'pwszPassword' )),
        COMMETHOD([], HRESULT, 'GetAccountInformation',
                  ( ['out'], POINTER(WSTRING), 'ppwszAccountName' )),
    ]

################################################################
##class ITask(object):
##    def GetApplicationName(self):
##        '-no docstring-'
##        #return ppwszApplicationName
##
##    def GetMaxRunTime(self):
##        '-no docstring-'
##        #return pdwMaxRunTimeMS
##
##    def SetWorkingDirectory(self, pwszWorkingDirectory):
##        '-no docstring-'
##        #return 
##
##    def SetParameters(self, pwszParameters):
##        '-no docstring-'
##        #return 
##
##    def SetTaskFlags(self, dwFlags):
##        '-no docstring-'
##        #return 
##
##    def GetWorkingDirectory(self):
##        '-no docstring-'
##        #return ppwszWorkingDirectory
##
##    def GetPriority(self):
##        '-no docstring-'
##        #return pdwPriority
##
##    def GetTaskFlags(self):
##        '-no docstring-'
##        #return pdwFlags
##
##    def SetApplicationName(self, pwszApplicationName):
##        '-no docstring-'
##        #return 
##
##    def SetPriority(self, dwPriority):
##        '-no docstring-'
##        #return 
##
##    def SetMaxRunTime(self, dwMaxRunTimeMS):
##        '-no docstring-'
##        #return 
##
##    def GetParameters(self):
##        '-no docstring-'
##        #return ppwszParameters
##
class ITask(IScheduledWorkItem):
    _case_insensitive_ = True
    _iid_ = GUID('{148BD524-A2AB-11CE-B11F-00AA00530503}')
    _idlflags_ = []
    _methods_ = [
        COMMETHOD([], HRESULT, 'SetApplicationName',
                  ( ['in'], WSTRING, 'pwszApplicationName' )),
        COMMETHOD([], HRESULT, 'GetApplicationName',
                  ( ['out'], POINTER(WSTRING), 'ppwszApplicationName' )),
        COMMETHOD([], HRESULT, 'SetParameters',
                  ( ['in'], WSTRING, 'pwszParameters' )),
        COMMETHOD([], HRESULT, 'GetParameters',
                  ( ['out'], POINTER(WSTRING), 'ppwszParameters' )),
        COMMETHOD([], HRESULT, 'SetWorkingDirectory',
                  ( ['in'], WSTRING, 'pwszWorkingDirectory' )),
        COMMETHOD([], HRESULT, 'GetWorkingDirectory',
                  ( ['out'], POINTER(WSTRING), 'ppwszWorkingDirectory' )),
        COMMETHOD([], HRESULT, 'SetPriority',
                  ( ['in'], c_ulong, 'dwPriority' )),
        COMMETHOD([], HRESULT, 'GetPriority',
                  ( ['out'], POINTER(c_ulong), 'pdwPriority' )),
        COMMETHOD([], HRESULT, 'SetTaskFlags',
                  ( ['in'], c_ulong, 'dwFlags' )),
        COMMETHOD([], HRESULT, 'GetTaskFlags',
                  ( ['out'], POINTER(c_ulong), 'pdwFlags' )),
        COMMETHOD([], HRESULT, 'SetMaxRunTime',
                  ( ['in'], c_ulong, 'dwMaxRunTimeMS' )),
        COMMETHOD([], HRESULT, 'GetMaxRunTime',
                  ( ['out'], POINTER(c_ulong), 'pdwMaxRunTimeMS' )),
    ]

################################################################
##class IEnumWorkItems(object):
##    def Reset(self):
##        '-no docstring-'
##        #return 
##
##    def Skip(self, celt):
##        '-no docstring-'
##        #return 
##
##    def Clone(self):
##        '-no docstring-'
##        #return ppEnumWorkItems
##
##    def Next(self, celt):
##        '-no docstring-'
##        #return rgpwszNames, pceltFetched
##
class IEnumWorkItems(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{148BD528-A2AB-11CE-B11F-00AA00530503}')
    _idlflags_ = []
    def __iter__(self):
        return self

    def next(self):
        item, fetched = self.Next(1)
        if fetched:
            return item
        raise StopIteration

    def __getitem__(self, index):
        self.Reset()
        self.Skip(index)
        item, fetched = self.Next(1)
        if fetched:
            return item
        raise IndexError(index)

IEnumWorkItems._methods_ = [
    COMMETHOD([], HRESULT, 'Next',
              ( ['in'], c_ulong, 'celt' ),
              ( ['out'], POINTER(POINTER(WSTRING)), 'rgpwszNames' ),
              ( ['out'], POINTER(c_ulong), 'pceltFetched' )),
    COMMETHOD([], HRESULT, 'Skip',
              ( ['in'], c_ulong, 'celt' )),
    COMMETHOD([], HRESULT, 'Reset'),
    COMMETHOD([], HRESULT, 'Clone',
              ( ['out'], POINTER(POINTER(IEnumWorkItems)), 'ppEnumWorkItems' )),
]

################################################################
##class ITaskScheduler(object):
##    def Activate(self, pwszName, riid):
##        '-no docstring-'
##        #return ppUnk
##
##    def NewWorkItem(self, pwszTaskName, rclsid, riid):
##        '-no docstring-'
##        #return ppUnk
##
##    def Enum(self):
##        '-no docstring-'
##        #return ppEnumWorkItems
##
##    def SetTargetComputer(self, pwszComputer):
##        '-no docstring-'
##        #return 
##
##    def IsOfType(self, pwszName, riid):
##        '-no docstring-'
##        #return 
##
##    def AddWorkItem(self, pwszTaskName, pWorkItem):
##        '-no docstring-'
##        #return 
##
##    def GetTargetComputer(self):
##        '-no docstring-'
##        #return ppwszComputer
##
##    def Delete(self, pwszName):
##        '-no docstring-'
##        #return 
##
class ITaskScheduler(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{148BD527-A2AB-11CE-B11F-00AA00530503}')
    _idlflags_ = []
    _methods_ = [
        COMMETHOD([], HRESULT, 'SetTargetComputer',
                  ( ['in'], WSTRING, 'pwszComputer' )),
        COMMETHOD([], HRESULT, 'GetTargetComputer',
                  ( ['out'], POINTER(WSTRING), 'ppwszComputer' )),
        COMMETHOD([], HRESULT, 'Enum',
                  ( ['out'], POINTER(POINTER(IEnumWorkItems)), 'ppEnumWorkItems' )),
        COMMETHOD([], HRESULT, 'Activate',
                  ( ['in'], WSTRING, 'pwszName' ),
                  ( ['in'], POINTER(GUID), 'riid' ),
                  ( ['out'], POINTER(POINTER(IUnknown)), 'ppUnk' )),
        COMMETHOD([], HRESULT, 'Delete',
                  ( ['in'], WSTRING, 'pwszName' )),
        COMMETHOD([], HRESULT, 'NewWorkItem',
                  ( ['in'], WSTRING, 'pwszTaskName' ),
                  ( ['in'], POINTER(GUID), 'rclsid' ),
                  ( ['in'], POINTER(GUID), 'riid' ),
                  ( ['out'], POINTER(POINTER(IUnknown)), 'ppUnk' )),
        COMMETHOD([], HRESULT, 'AddWorkItem',
                  ( ['in'], WSTRING, 'pwszTaskName' ),
                  ( ['in'], POINTER(IScheduledWorkItem), 'pWorkItem' )),
        COMMETHOD([], HRESULT, 'IsOfType',
                  ( ['in'], WSTRING, 'pwszName' ),
                  ( ['in'], POINTER(GUID), 'riid' )),
    ]

################################################################
### COM Classes
################################################################

class CTask(CoClass):
    _reg_clsid_ = GUID('{148BD520-A2AB-11CE-B11F-00AA00530503}')
    _idlflags_ = []
    _typelib_path_ = typelib_path
    _com_interfaces_ = [ITask]

class CTaskScheduler(CoClass):
    _reg_clsid_ = GUID('{148BD52A-A2AB-11CE-B11F-00AA00530503}')
    _idlflags_ = []
    _typelib_path_ = typelib_path
    _com_interfaces_ = [ITaskScheduler]


__all__ = ['NORMAL_PRIORITY_CLASS',
           'IDLE_PRIORITY_CLASS',
           'HIGH_PRIORITY_CLASS',
           'REALTIME_PRIORITY_CLASS',
           'SCHED_S_TASK_READY',
           'SCHED_S_TASK_RUNNING',
           'SCHED_S_TASK_DISABLED',
           'SCHED_S_TASK_HAS_NOT_RUN',
           'SCHED_S_TASK_NO_MORE_RUNS',
           'SCHED_S_TASK_NOT_SCHEDULED',
           'SCHED_S_TASK_TERMINATED',
           'SCHED_S_TASK_NO_VALID_TRIGGERS',
           'SCHED_S_EVENT_TRIGGER',
           'SCHED_E_TRIGGER_NOT_FOUND',
           'SCHED_E_TASK_NOT_READY',
           'SCHED_E_TASK_NOT_RUNNING',
           'SCHED_E_SERVICE_NOT_INSTALLED',
           'SCHED_E_CANNOT_OPEN_TASK',
           'SCHED_E_INVALID_TASK',
           'SCHED_E_ACCOUNT_INFORMATION_NOT_SET',
           'SCHED_E_ACCOUNT_NAME_NOT_FOUND',
           'SCHED_E_ACCOUNT_DBASE_CORRUPT',
           'SCHED_E_NO_SECURITY_SERVICES',
           'SCHED_E_UNKNOWN_OBJECT_VERSION',
           'SCHED_E_UNSUPPORTED_ACCOUNT_OPTION',
           'SCHED_E_SERVICE_NOT_RUNNING',
           'SCHED_E_UNEXPECTEDNODE',
           'SCHED_E_NAMESPACE',
           'SCHED_E_INVALIDVALUE',
           'SCHED_E_MISSINGNODE',
           'SCHED_E_MALFORMEDXML',
           'SCHED_S_SOME_TRIGGERS_FAILED',
           'SCHED_S_BATCH_LOGON_PROBLEM',
           'SCHED_E_TOO_MANY_NODES',
           'SCHED_E_PAST_END_BOUNDARY',
           'SCHED_E_ALREADY_RUNNING',
           'SCHED_E_USER_NOT_LOGGED_ON',
           'SCHED_E_INVALID_TASK_HASH',
           'SCHED_E_SERVICE_NOT_AVAILABLE',
           'SCHED_E_SERVICE_TOO_BUSY',
           'SCHED_E_TASK_ATTEMPTED',
           'SCHED_S_TASK_QUEUED',
           'SCHED_E_TASK_DISABLED',
           'SCHED_E_TASK_NOT_V1_COMPAT',
           'SCHED_E_START_ON_DEMAND',
           'TASK_SUNDAY',
           'TASK_MONDAY',
           'TASK_TUESDAY',
           'TASK_WEDNESDAY',
           'TASK_THURSDAY',
           'TASK_FRIDAY',
           'TASK_SATURDAY',
           'TASK_JANUARY',
           'TASK_FEBRUARY',
           'TASK_MARCH',
           'TASK_APRIL',
           'TASK_MAY',
           'TASK_JUNE',
           'TASK_JULY',
           'TASK_AUGUST',
           'TASK_SEPTEMBER',
           'TASK_OCTOBER',
           'TASK_NOVEMBER',
           'TASK_DECEMBER',
           'TASK_FIRST_WEEK',
           'TASK_SECOND_WEEK',
           'TASK_THIRD_WEEK',
           'TASK_FOURTH_WEEK',
           'TASK_LAST_WEEK',
           'TASK_FLAG_INTERACTIVE',
           'TASK_FLAG_DELETE_WHEN_DONE',
           'TASK_FLAG_DISABLED',
           'TASK_FLAG_START_ONLY_IF_IDLE',
           'TASK_FLAG_KILL_ON_IDLE_END',
           'TASK_FLAG_DONT_START_IF_ON_BATTERIES',
           'TASK_FLAG_KILL_IF_GOING_ON_BATTERIES',
           'TASK_FLAG_RUN_ONLY_IF_DOCKED',
           'TASK_FLAG_HIDDEN',
           'TASK_FLAG_RUN_IF_CONNECTED_TO_INTERNET',
           'TASK_FLAG_RESTART_ON_IDLE_RESUME',
           'TASK_FLAG_SYSTEM_REQUIRED',
           'TASK_FLAG_RUN_ONLY_IF_LOGGED_ON',
           'TASK_TRIGGER_FLAG_HAS_END_DATE',
           'TASK_TRIGGER_FLAG_KILL_AT_DURATION_END',
           'TASK_TRIGGER_FLAG_DISABLED',
           'TASK_MAX_RUN_TIMES',
           'TASK_TIME_TRIGGER_ONCE',
           'TASK_TIME_TRIGGER_DAILY',
           'TASK_TIME_TRIGGER_WEEKLY',
           'TASK_TIME_TRIGGER_MONTHLYDATE',
           'TASK_TIME_TRIGGER_MONTHLYDOW',
           'TASK_EVENT_TRIGGER_ON_IDLE',
           'TASK_EVENT_TRIGGER_AT_SYSTEMSTART',
           'TASK_EVENT_TRIGGER_AT_LOGON',
           '_TASK_TRIGGER',
           '_TASK_TRIGGER_TYPE',
           '_TRIGGER_TYPE_UNION',
           '_DAILY',
           '_WEEKLY',
           '_MONTHLYDATE',
           '_MONTHLYDOW',
           '_SYSTEMTIME',
           'ITaskScheduler',
           'ITask',
           'IScheduledWorkItem',
           'IEnumWorkItems',
           'ITaskTrigger',
           'CTaskScheduler',
           'CTask']
from comtypes import _check_version; _check_version('501')
