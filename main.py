import os
from kompas_connection import get_kompas_api7
from json_runner import JsonKompasRunner

def main():
    # json_file = "jsons_files\\test.json"
    json_file = "jsons_files\\opore.json"
    # json_file = "jsons_files\\models_m3d.json"
    if not os.path.exists(json_file):
        print(f"Файл {json_file} не найден!")
        return

    # 1. Подключаемся к КОМПАСу
    print("Подключение к КОМПАС-3D...")
    module7, app, const7 = get_kompas_api7()
    app.Visible = True  # Делаем компас видимым

    # 2. Инициализируем раннер
    runner = JsonKompasRunner(module7, app, const7)

    # 3. Выполняем файл
    runner.execute_script(json_file)

if __name__ == "__main__":
    main()
