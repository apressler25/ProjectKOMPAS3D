import json
import operations

import json
import inspect
import operations

class JsonKompasRunner:
    def __init__(self, module7, app, const7):
        self.module7 = module7
        self.app = app
        self.const7 = const7
        self.ctx = {}

        self.operations_map = {}
        for name, func in inspect.getmembers(operations, inspect.isfunction):
            if name.startswith("op_"):
                json_command_name = name[3:] 
                self.operations_map[json_command_name] = func


    def execute_script(self, json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            script = json.load(f)

        print(f"--- Запуск скрипта: версия {script.get('version', '1.0')} ---")
        
        for step in script.get("operations", []):
            op_name = step.get("op")
            params = step.get("params", {})
            
            # Ищем функцию в словаре по ключу (op_name)
            handler = self.operations_map.get(op_name)
            if not handler:
                print(f"ОШИБКА: Неизвестная операция '{op_name}'")
                continue
                
            # Запускаем функцию
            handler(self, params)
        print("--- Выполнение завершено ---")
