import json
import pythoncom
from pathlib import Path
from kompas_connection import get_kompas_api7

class KompasM3DParser:
    def __init__(self):
        print("Подключение к КОМПАС-3D...")
        self.module7, self.app, self.const7 = get_kompas_api7()
        self.app.Visible = False

    def parse_directory(self, target_dir_path):
        """
        Парсит все .m3d файлы в указанной директории.
        :param target_dir_path: Абсолютный или относительный путь к папке.
        """
        target_dir = Path(target_dir_path)

        if not target_dir.exists():
            print(f"[-] Директория '{target_dir}' не найдена!")
            return {}

        results = {}
        m3d_files = list(target_dir.glob("*.m3d"))
        print(f"Найдено моделей: {len(m3d_files)}")

        for file_path in m3d_files:
            print(f"Парсинг: {file_path.name}...")
            parsed_data = self.parse_model(str(file_path))
            if parsed_data:
                results[file_path.name] = parsed_data

        return results

    def parse_model(self, filepath: str):
        doc = self.app.Documents.Open(filepath, False, True)
        if not doc:
            print(f" [-] Не удалось открыть: {filepath}")
            return None

        try:
            doc3d = self.module7.IKompasDocument3D(doc)
            top_part = doc3d.TopPart
            
            model_data = {
                "PartVariables": self._get_variables(top_part),
                "OperationsTree": self._parse_features_tree(top_part)
            }
            return model_data

        finally:
            doc.Close(self.const7.kdDoNotSaveChanges)

    def _parse_features_tree(self, feature_obj):
        tree = []
        try:
            feature = self.module7.IFeature7(feature_obj)
            sub_features = feature.SubFeatures(self.const7.ksOperTree, True, False)

            if sub_features:
                try:
                    count = sub_features.Count
                    iterator = range(count)
                except AttributeError:
                    iterator = range(len(sub_features))

                for idx, sub_feat in enumerate(sub_features):
                    sub_feat7 = self.module7.IFeature7(sub_feat)

                    ref_id = None
                    try:
                        ref_id = sub_feat7.Reference
                    except AttributeError:
                        pass
                        
                    feat_data = {
                        "IndexInTree": idx,  
                        "ReferenceID": ref_id, 
                        "ActionName": sub_feat7.Name,
                        "FeatureType": sub_feat7.FeatureType,
                        "Variables": self._get_variables(sub_feat7),
                        "Geometry": self._get_sketch_geometry(sub_feat7),
                        "SubActions": self._parse_features_tree(sub_feat7)
                    }
                    tree.append(feat_data)
        except Exception:
            pass
            
        return tree

    def _get_sketch_geometry(self, feature_obj):
        """Извлекает 2D-геометрию (координаты линий) через прямой запрос интерфейса ISketch."""
        geometry_data = {"Lines": []}

        try:
            sketch_disp = feature_obj._oleobj_.QueryInterface(self.module7.ISketch.CLSID, pythoncom.IID_IDispatch)
            sketch = self.module7.ISketch(sketch_disp)

            doc2d_obj = sketch.BeginEdit()
            
            if doc2d_obj:
                doc2d = self.module7.IKompasDocument2D(doc2d_obj)
                active_view = doc2d.ViewsAndLayersManager.Views.ActiveView
                container = self.module7.IDrawingContainer(active_view)

                lines = container.LineSegments
                if lines:
                    if hasattr(lines, 'Count'):
                        for i in range(lines.Count):
                            line = self.module7.ILineSegment(lines.Item(i))
                            geometry_data["Lines"].append({
                                "X1": round(line.X1, 3),
                                "Y1": round(line.Y1, 3),
                                "X2": round(line.X2, 3),
                                "Y2": round(line.Y2, 3)
                            })
                    else:
                        for line_obj in lines:
                            line = self.module7.ILineSegment(line_obj)
                            geometry_data["Lines"].append({
                                "X1": round(line.X1, 3),
                                "Y1": round(line.Y1, 3),
                                "X2": round(line.X2, 3),
                                "Y2": round(line.Y2, 3)
                            })
                
                circles = container.Circles
                if circles:
                    if not "Circles" in geometry_data: geometry_data["Circles"] = []
                    if hasattr(circles, 'Count'):
                        for i in range(circles.Count):
                            circ = self.module7.ICircle(circles.Item(i))
                            geometry_data["Circles"].append({
                                "Xc": round(circ.Xc, 3),
                                "Yc": round(circ.Yc, 3),
                                "Radius": round(circ.Radius, 3)
                            })
                    else:
                        for circ_obj in circles:
                            circ = self.module7.ICircle(circ_obj)
                            geometry_data["Circles"].append({
                                "Xc": round(circ.Xc, 3),
                                "Yc": round(circ.Yc, 3),
                                "Radius": round(circ.Radius, 3)
                            })

        except Exception as e:
            return None
            
        finally:
            try:
                if 'sketch' in locals():
                    sketch.EndEdit()
            except:
                pass

        return geometry_data if (geometry_data.get("Lines") or geometry_data.get("Circles")) else None

    def _get_variables(self, feature_obj):
        vars_dict = {}
        try:
            feature = self.module7.IFeature7(feature_obj)
            variables_col = feature.Variables(False, False)

            if variables_col:
                try:
                    count = variables_col.Count
                    for i in range(count):
                        var7 = self.module7.IVariable7(variables_col.Item(i))
                        vars_dict[var7.Name] = {"Value": var7.Value, "Expression": var7.Expression}
                except AttributeError:
                    for var in variables_col:
                        var7 = self.module7.IVariable7(var)
                        vars_dict[var7.Name] = {"Value": var7.Value, "Expression": var7.Expression}
        except Exception:
            pass

        return vars_dict

if __name__ == "__main__":
    parser = KompasM3DParser()
    test_dir = Path.cwd() / "models_kompas"
    
    parsed_data = parser.parse_directory(str(test_dir))

    output_file = "kompas_history_geometry_dump.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(parsed_data, f, ensure_ascii=False, indent=4)

    print(f"\nГотово! Результаты сохранены в файл: {output_file}")