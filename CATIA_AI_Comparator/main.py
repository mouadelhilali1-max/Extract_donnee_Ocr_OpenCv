from catia.connection import CATIAConnection
from catia.visibility_manager import VisibilityManager

catia = CATIAConnection()

if catia.connect():

    vm = VisibilityManager(catia.document)

    names = [
        "PREPARATIONS",
        "CPC",
        "OFFSET DECAL IT",
        "FOR INFORMATION"
    ]

    for name in names:

        obj = vm.find_hybrid_body(name)

        if obj:

            vm.set_visibility(obj, False)

        else:

            print(f"{name} introuvable")