import os
from kompas_connection import get_kompas_api7
from json_runner import JsonKompasRunner

def main():
    json_file = "jsons_files\\opore.json"
    if not os.path.exists(json_file):
        print(f"Файл {json_file} не найден!")
        return

    print("Подключение к КОМПАС-3D...")
    module7, app, const7 = get_kompas_api7()
    app.Visible = True 

    runner = JsonKompasRunner(module7, app, const7)

    runner.execute_script(json_file)

if __name__ == "__main__":
    main()
