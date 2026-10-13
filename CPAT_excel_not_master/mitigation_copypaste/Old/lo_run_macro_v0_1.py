"""Run a VBA macro of an .xlsm in headless LibreOffice (VBA compatibility mode) and save the result as .xlsx.

Used by check_v1_03.py to test the embedded module CPATScenarios: LibreOffice imports the VBA project of the file
(library VBAProject), runs the macro with Application.Calculate recalculating the workbook, and the saved values
are compared with the Python emulation. LibreOffice cannot evaluate LAMBDA, so the test file has its LAMBDA column
expanded first.

    python lo_run_macro_v0_1.py in.xlsm Module.Procedure out.xlsx
"""
import os
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def main():
    src, macro, dst = sys.argv[1:4]
    prof = tempfile.mkdtemp(prefix='lo_macro_profile_')
    port = 2100 + os.getpid() % 800
    proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', f'-env:UserInstallation=file://{prof}',
                             f'--accept=socket,host=localhost,port={port};urp;'],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    ctx = None
    for _ in range(120):
        try:
            local = uno.getComponentContext()
            res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
            ctx = res.resolve(f'uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext')
            break
        except Exception:
            time.sleep(1)
    smgr = ctx.ServiceManager
    desk = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
    for name in ('Load', 'Executable'):                    # import VBA code and make it executable
        acc = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess',
                                             (pv('nodepath', '/org.openoffice.Office.Calc/Filter/Import/VBA'),))
        acc.setPropertyValue(name, True)
        acc.commitChanges()
    try:
        doc = desk.loadComponentFromURL('file://' + os.path.abspath(src), '_blank', 0,
                                        (pv('Hidden', True), pv('MacroExecutionMode', 4)))
        mod, sub = macro.split('.')
        libs, lib = doc.BasicLibraries, None
        for ln in libs.getElementNames():
            libs.loadLibrary(ln)
            if mod in libs.getByName(ln).getElementNames():
                lib = ln
        assert lib, f'module {mod} not found'
        t = time.time()
        script = doc.getScriptProvider().getScript(
            f'vnd.sun.star.script:{lib}.{mod}.{sub}?language=Basic&location=document')
        result = script.invoke((), (), ())
        print(f'{lib}.{mod}.{sub} returned {result[0]!r} in {time.time() - t:.1f} s')
        doc.storeToURL('file://' + os.path.abspath(dst), (pv('FilterName', 'Calc MS Excel 2007 XML'),))
        doc.close(True)
    finally:
        try:
            desk.terminate()
        except Exception:
            pass
        proc.wait(timeout=120)


if __name__ == '__main__':
    main()
