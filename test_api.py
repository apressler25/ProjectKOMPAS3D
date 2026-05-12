import pythoncom
from win32com.client import Dispatch, gencache

def get_kompas_api7():
    """
    Функция для подключения к API версии 7 программы КОМПАС-3D.
    """
    # Уникальные идентификаторы (GUID) для модулей API7
    # Эти GUID можно найти в документации SDK или в системном реестре
    KOMPAS_API7_GUID = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"
    KOMPAS_CONSTANTS_GUID = "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}"

    # Создаем или получаем модули для работы с API7 и константами
    module7 = gencache.EnsureModule(KOMPAS_API7_GUID, 0, 1, 0)
    const7 = gencache.EnsureModule(KOMPAS_CONSTANTS_GUID, 0, 1, 0).constants

    # Создаем COM-объект КОМПАС-3D (используется старый ProgID 'Kompas.Application.7')
    # Даже для новых версий Компаса этот ProgID работает и возвращает объект API5,
    # из которого мы затем получим API7.
    kompas_object = Dispatch("Kompas.Application.7")

    # Запрашиваем у объекта API5 интерфейс API7
    # Для этого используется QueryInterface с IID_IDispatch
    api7 = module7.IKompasAPIObject(
        kompas_object._oleobj_.QueryInterface(module7.IKompasAPIObject.CLSID,
                                            pythoncom.IID_IDispatch)
    )
    return module7, api7, const7

# --- Основная часть программы ---
if __name__ == "__main__":
    # Шаг 1: Подключаемся к API7
    print("Подключаемся к КОМПАС-3D...")
    module7, api7, const7 = get_kompas_api7()

    # Шаг 2: Получаем главный интерфейс приложения IApplication
    app7 = api7.Application

    # Шаг 3: Делаем окно КОМПАС-3D видимым (если оно было скрыто)
    # Если КОМПАС еще не запущен, эта команда его запустит.
    app7.Visible = True
    print("Окно КОМПАС-3D открыто.")

    # Шаг 4: Подавляем появление диалоговых окон с вопросами к пользователю.
    # Это полезно для автоматических скриптов.
    # ksHideMessageNo означает, что на любой вопрос программа "ответит" НЕТ.
    app7.HideMessage = const7.ksHideMessageNo
    print("Диалоговые окна подавлены.")

    # Шаг 5: Получаем и выводим полное название программы
    # Метод ApplicationName возвращает строку с названием.
    app_name = app7.ApplicationName(FullName=True)
    print(f"Успешно подключились к: {app_name}")

    # Здесь можно будет добавить другой функционал, например, создание документа.
    # input("Нажмите Enter для завершения...") # Чтобы окно не закрылось сразу

    print("Программа завершена.")