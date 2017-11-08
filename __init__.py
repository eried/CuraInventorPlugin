# Copyright (c) 2017 Thomas Karl Pietrowski

from UM.Platform import Platform

from UM.i18n import i18nCatalog
i18n_catalog = i18nCatalog("InventorPlugin")

if Platform.isWindows():
    # For installation check
    import winreg
    # The reader plugin itself
    from . import InventorReader


def getMetaData():
    metaData = {"mesh_reader": [],
                }
    
    if InventorReader.is_askinv_service():
        metaData["mesh_reader"] += [{
                                        "extension": "IPT",
                                        "description": i18n_catalog.i18nc("@item:inlistbox", "Inventor part file")
                                    },
                                    {
                                        "extension": "IAM",
                                        "description": i18n_catalog.i18nc("@item:inlistbox", "Inventor assembly file")
                                    },
                                    {
                                        "extension": "DWG",
                                        "description": i18n_catalog.i18nc("@item:inlistbox", "Inventor drawing file")
                                    },
                                    ]
    
    return metaData

def register(app):
    # Solid works only runs on Windows.
    plugin_data = {}
    if Platform.isWindows():
        reader = InventorReader.InventorReader()
        plugin_data["mesh_reader"] = reader
    return plugin_data
