import pythoncom
from win32com.client import Dispatch, gencache

def get_kompas_api7():
    """
    Функция для подключения к API версии 7 программы КОМПАС-3D.
    """
    KOMPAS_API7_GUID = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"
    KOMPAS_CONSTANTS_GUID = "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}"

    module7 = gencache.EnsureModule(KOMPAS_API7_GUID, 0, 1, 0)
    const7 = gencache.EnsureModule(KOMPAS_CONSTANTS_GUID, 0, 1, 0).constants

    kompas_object = Dispatch("Kompas.Application.7")

    api7 = module7.IKompasAPIObject(
        kompas_object._oleobj_.QueryInterface(module7.IKompasAPIObject.CLSID,
                                              pythoncom.IID_IDispatch)
    )
    
    # Получаем главный интерфейс приложения IApplication из базового
    app = module7.IApplication(api7)
    
    return module7, app, const7
