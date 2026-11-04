import sys
sys.stdout.reconfigure(encoding="utf-8")
import jpype, mpxj
jpype.startJVM()
from org.mpxj.reader import UniversalProjectReader
p = UniversalProjectReader().read(sys.argv[1])
print("tasks", p.getTasks().size(), "resources", p.getResources().size(), "assignments", p.getResourceAssignments().size())
cal = p.getDefaultCalendar()
print("calendar", cal.getName(), [str(e.getName()) + " " + str(e.getFromDate()) + " working=" + str(e.getWorking()) for e in cal.getCalendarExceptions()])
for r in p.getResources():
    if r.getName():
        print("res", r.getName(), r.getStandardRate())
tot = 0
for t in p.getTasks():
    if t.getWBS() in ("1.2.7", "1.5.3", "1.6.6", "1.7.4", "1.7.4.1", "1.7.5.8", "1.5.1"):
        print(t.getWBS(), t.getName(), t.getStart(), t.getFinish(), "fixed", t.getFixedCost(), "constraint", t.getConstraintType(), t.getConstraintDate(), "summary", t.getSummary())
jpype.shutdownJVM()
