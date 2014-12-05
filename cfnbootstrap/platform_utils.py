import logging
import sys

_task_name = "cfn-init Resume Trigger"

_scheduler_supported = True
try:
    import comtypes.client
    from _ctypes import COMError
    from ctypes import sizeof
    from comtypes.GUID import GUID
    from comtypes.persist import IPersistFile
    from mstask import *
except ImportError:
    _scheduler_supported = False

log = logging.getLogger("cfn.init")

def set_reboot_trigger():
    if not _scheduler_supported:
        log.debug("Not setting a reboot trigger as scheduling support is not available")
        return

    scheduler = comtypes.client.CreateObject(CTaskScheduler, comtypes.CLSCTX_INPROC_SERVER, None, ITaskScheduler)

    log.debug("Creating Scheduled Task for cfn-init resume")
    
    try:
        task_definition = scheduler.NewWorkItem(_task_name, CTask._reg_clsid_, ITask._iid_)
    except COMError:
        log.debug("Scheduled task already exists; not updating")
        return
    
    task_definition = task_definition.QueryInterface(ITask)
    
    task_definition.SetComment("This will resume cfn-init after the system has booted")
    task_definition.SetCreator("Amazon Web Services")
    
    task_definition.SetFlags(TASK_FLAG_RUN_IF_CONNECTED_TO_INTERNET)
    #task_definition.SetErrorRetryCount(3)    # not implemented
    #task_definition.SetErrorRetryInterval(1) # not implemented
    
    trigger = task_definition.CreateTrigger()[1]
    Trigger = _TASK_TRIGGER()
    Trigger.cbTriggerSize = sizeof(_TASK_TRIGGER)
    Trigger.wBeginDay     = 1                                 # Required
    Trigger.wBeginMonth   = 1                                 # Required
    Trigger.wBeginYear    = 1999                              # Required
    Trigger.TriggerType   = TASK_EVENT_TRIGGER_AT_SYSTEMSTART
    
    trigger.SetTrigger(Trigger)
    #trigger.Id = "CfnInitBootTrigger"
    #trigger.Delay = "PT30S"
    
    task_definition.SetApplicationName(sys.executable)
    task_definition.SetParameters("-v --resume")
    
    task_definition.SetAccountinformation("", None)
    task_definition.SetPriority(256)
    
    persist_file = task_definition.QueryInterface(IPersistFile)
    persist_file.Save(None, True)
    
    log.debug("Scheduled Task created")

def clear_reboot_trigger():
    if not _scheduler_supported:
        log.debug("Not clearing reboot trigger as scheduling support is not available")
        return
    
    scheduler = comtypes.client.CreateObject(CTaskScheduler, comtypes.CLSCTX_INPROC_SERVER, None, ITaskScheduler)
    
    log.debug("Deleting Scheduled Task for cfn-init resume")
    
    try:
        scheduler.Delete(_task_name)
        log.debug("Scheduled Task deleted")
    except com_error:
        log.debug("Scheduled Task did not exist")

