"""
Модуль содержит операции для параметрического моделирования в КОМПАС-3D.
Каждая операция (op_*) принимает объект runner и словарь параметров (params).
Большинство операций поддерживают сохранение своего результата по имени (save_as) 
и использование ранее сохраненных результатов (use_feature).
"""

def op_create_document_3d(runner, params):
    """
    Создает новый 3D документ (Деталь).
    
    Параметры (params):
      - visible (bool, по умолчанию True): делать ли документ видимым на экране.
    """
    # const7.ksDocumentPart == 4 (Деталь)
    visible = params.get("visible", True)
    doc3d = runner.app.Documents.Add(runner.const7.ksDocumentPart, visible)
    
    # Сохраняем в контекст интерфейс IPart7, чтобы строить на нем эскизы
    ik_doc3d = runner.module7.IKompasDocument3D(doc3d)
    runner.ctx["part"] = ik_doc3d.TopPart
    print("Создан 3D документ")


# def op_create_sketch_on_plane(runner, params):
#     """
#     Создает эскиз на базовой плоскости (XY, ZX, YOZ).
    
#     Параметры (params):
#       - plane (str, по умолчанию "XY"): имя плоскости ("XY", "ZX", "YZ").
#       - save_as (str, опционально): имя, под которым сохранить эскиз в контексте.
#     """
#     plane_str = params.get("plane", "XY")
#     save_as = params.get("save_as")
#     part = runner.ctx["part"]
    
#     # Прямые числовые коды констант ksObj3dTypeEnum:
#     if plane_str == "XY":
#         plane = part.DefaultObject(1)  # 1 = XOY
#     elif plane_str == "ZX":
#         plane = part.DefaultObject(2)  # 2 = XOZ
#     else:
#         plane = part.DefaultObject(3)  # 3 = YOZ
        
#     model_container = runner.module7.IModelContainer(part)
#     sketch = model_container.Sketchs.Add()
#     sketch.Plane = plane
#     sketch.Update()
    
#     # Входим в режим редактирования эскиза
#     sketch.BeginEdit()
    
#     runner.ctx["sketch"] = sketch
#     if save_as:
#         if "saved_features" not in runner.ctx:
#             runner.ctx["saved_features"] = {}
#         runner.ctx["saved_features"][save_as] = sketch
        
#     print(f"Создан эскиз на плоскости {plane_str}" + (f" (сохранен как '{save_as}')" if save_as else ""))


def op_extrude_sketch(runner, params):
    """
    Выдавливает текущий эскиз, добавляя материал (Boss Extrusion).
    
    Параметры (params):
      - length (float, по умолчанию 50.0): длина выдавливания в мм.
      - save_as (str, опционально): имя, под которым сохранить операцию в контексте.
    """
    part = runner.ctx["part"]
    sketch = runner.ctx["sketch"]
    length = params.get("length", 50.0)
    save_as = params.get("save_as")
    
    model_container = runner.module7.IModelContainer(part)
    extrusions = model_container.Extrusions
    
    # 24 = o3d_bossExtrusion
    extrusion = extrusions.Add(24)
    
    if extrusion is None:
        raise RuntimeError("Не удалось создать операцию выдавливания.")
        
    import win32com.client
    try:
        extrusion1 = win32com.client.CastTo(extrusion, 'IExtrusion1')
        extrusion1.Profile = sketch
    except Exception:
        extrusion.Profile = sketch
    
    # 0 = dtNormal (Прямое направление)
    extrusion.Direction = 0
    extrusion.SetSideParameters(True, 0, length, 0, False, None)
    extrusion.Update()
    
    runner.ctx["last_feature"] = extrusion
    if save_as:
        if "saved_features" not in runner.ctx:
            runner.ctx["saved_features"] = {}
        runner.ctx["saved_features"][save_as] = extrusion
        
    print(f"Выполнено выдавливание на {length} мм" + (f" (сохранено как '{save_as}')" if save_as else ""))


def op_extrude_cut(runner, params):
    """
    Выдавливает текущий эскиз с вырезанием материала (Cut Extrusion).
    Идеально для создания отверстий.
    
    Параметры (params):
      - length (float, по умолчанию 10.0): глубина выреза.
      - through_all (bool, по умолчанию False): вырезать насквозь через всю деталь.
      - direction (str, по умолчанию "normal"): направление выреза ("normal" или "reverse").
      - operation_result (int, по умолчанию 2): результат операции (1 - объединение, 2 - вычитание, 3 - пересечение, 4 - новое тело).
      - save_as (str, опционально): имя для сохранения операции.
    """
    part = runner.ctx["part"]
    sketch = runner.ctx["sketch"]
    length = params.get("length", 10.0)
    through_all = params.get("through_all", False)
    dir_str = params.get("direction", "normal").lower()
    
    # 2 = ksORSubtraction (Вычитание)
    operation_result = params.get("operation_result", 2) 
    save_as = params.get("save_as")
    
    model_container = runner.module7.IModelContainer(part)
    extrusions = model_container.Extrusions
    
    # 25 = o3d_cutExtrusion (Вырезать выдавливанием)
    extrusion = extrusions.Add(25)
    
    import win32com.client
    # Чтобы задать результат операции (Вычитание) и эскиз, нужен интерфейс IExtrusion1
    try:
        extrusion1 = win32com.client.CastTo(extrusion, 'IExtrusion1')
        extrusion1.Profile = sketch
        extrusion1.OperationResult = operation_result # <--- Указываем, что это Вычитание!
    except Exception as e:
        print(f"Предупреждение: не удалось получить интерфейс IExtrusion1 ({e}). Пробуем через базовый.")
        extrusion.Profile = sketch
    
    # Если through_all = True, используем etThroughAll (1), иначе etBlind (0)
    end_type = 1 if through_all else 0
    
    # Настраиваем направление
    if dir_str == "reverse":
        extrusion.Direction = 1  # dtReverse
        extrusion.SetSideParameters(False, end_type, length, 0, False, None)
    else:
        extrusion.Direction = 0  # dtNormal
        extrusion.SetSideParameters(True, end_type, length, 0, False, None)
        
    extrusion.Update()
    
    runner.ctx["last_feature"] = extrusion
    if save_as:
        if "saved_features" not in runner.ctx:
            runner.ctx["saved_features"] = {}
        runner.ctx["saved_features"][save_as] = extrusion
        
    msg = "насквозь" if through_all else f"на {length} мм"
    dir_msg = " (в обратном направлении)" if dir_str == "reverse" else ""
    print(f"Выполнен вырез {msg}{dir_msg} (тип операции: {operation_result})" + (f" (сохранено как '{save_as}')" if save_as else ""))


def op_draw_rectangle_center(runner, params):
    """
    Рисует прямоугольник из центра в текущем эскизе.
    
    Параметры (params):
      - center (list[float], по умолчанию [0.0, 0.0]): координаты центра [X, Y].
      - width (float, по умолчанию 50.0): ширина прямоугольника.
      - height (float, по умолчанию 50.0): высота прямоугольника.
    """
    center_x, center_y = params.get("center", [0.0, 0.0])
    width = params.get("width", 50.0)
    height = params.get("height", 50.0)
    
    import win32com.client
    # Получаем активный документ (эскиз в режиме редактирования)
    app7 = win32com.client.Dispatch("Kompas.Application.7")
    sketch_doc = app7.ActiveDocument
    
    # Приводим к интерфейсу 2D-документа
    doc2d = win32com.client.CastTo(sketch_doc, 'IKompasDocument2D')
    
    # Получаем активный вид для рисования
    views = doc2d.ViewsAndLayersManager.Views
    active_view = views.ActiveView
    drawing_container = win32com.client.CastTo(active_view, 'IDrawingContainer')
    
    # Создаем прямоугольник
    rectangles = drawing_container.Rectangles
    rect = rectangles.Add()
    rect.X = center_x - width / 2
    rect.Y = center_y - height / 2
    rect.Width = width
    rect.Height = height
    rect.Update()
    
    print(f"Нарисован прямоугольник {width}x{height}")



def op_draw_circle(runner, params):
    """
    Рисует окружность в текущем эскизе.
    
    Параметры (params):
      - center (list[float], по умолчанию [0.0, 0.0]): координаты центра [X, Y].
      - radius (float, по умолчанию 20.0): радиус окружности.
    """
    center = params.get("center", [0.0, 0.0])
    x, y = center[0], center[1]
    radius = params.get("radius", 20.0)
    
    import win32com.client
    app7 = win32com.client.Dispatch("Kompas.Application.7")
    sketch_doc = runner.ctx.get("sketch_doc", app7.ActiveDocument)
    doc2d = win32com.client.CastTo(sketch_doc, 'IKompasDocument2D')
    view = doc2d.ViewsAndLayersManager.Views.ActiveView
    drawing_container = win32com.client.CastTo(view, 'IDrawingContainer')
    
    circles = drawing_container.Circles
    circle = circles.Add()
    circle.Xc = x
    circle.Yc = y
    circle.Radius = radius
    circle.Update()
    
    print(f"Нарисована окружность радиусом {radius} в точке ({x}, {y})")


def op_draw_polygon(runner, params):
    """
    Рисует правильный многоугольник в текущем эскизе.
    
    Параметры (params):
      - center (list[float], по умолчанию [0.0, 0.0]): координаты центра [X, Y].
      - sides (int, по умолчанию 6): количество сторон.
      - radius (float, по умолчанию 10.0): радиус описанной/вписанной окружности.
      - describe (bool, по умолчанию True): True - описанный, False - вписанный.
    """
    center = params.get("center", [0.0, 0.0])
    x, y = center[0], center[1]
    sides = params.get("sides", 6)
    radius = params.get("radius", 10.0)
    describe = params.get("describe", True)
    
    import win32com.client
    app7 = win32com.client.Dispatch("Kompas.Application.7")
    sketch_doc = runner.ctx.get("sketch_doc", app7.ActiveDocument)
    doc2d = win32com.client.CastTo(sketch_doc, 'IKompasDocument2D')
    view = doc2d.ViewsAndLayersManager.Views.ActiveView
    drawing_container = win32com.client.CastTo(view, 'IDrawingContainer')
    
    polygons = drawing_container.RegularPolygons
    polygon = polygons.Add()
    polygon.Count = sides
    polygon.Radius = radius
    polygon.Xc = x
    polygon.Yc = y
    polygon.Describe = describe 
    polygon.Update()
    
    print(f"Нарисован {sides}-угольник (R={radius}) в точке ({x}, {y})")


# def op_finish_sketch(runner, params):
#     """
#     Завершает редактирование текущего эскиза.
    
#     Параметры (params):
#       - save_as (str, опционально): имя для сохранения эскиза в словаре.
#     """
#     sketch = runner.ctx.get("sketch")
#     part = runner.ctx.get("part")
#     save_as = params.get("save_as")
    
#     if sketch:
#         sketch.EndEdit()
#         sketch.Update()
        
#         if part:
#             part.Update()
            
#         runner.ctx["last_feature"] = sketch
#         runner.ctx.pop("last_point", None)
#         runner.ctx.pop("first_point", None)
#         if save_as:
#             if "saved_features" not in runner.ctx:
#                 runner.ctx["saved_features"] = {}
#             runner.ctx["saved_features"][save_as] = sketch
            
#         print("Эскиз завершен" + (f" (сохранен как '{save_as}')" if save_as else ""))
#     else:
#         print("Нет активного эскиза для завершения")


def op_chamfer_all_edges(runner, params):
    """
    Строит фаску на всех ребрах указанной операции.
    
    Параметры (params):
      - distance (float, по умолчанию 1.0): размер катета фаски.
      - use_feature (str, опционально): имя операции (например выдавливания), ребра которой нужно скруглить. Если не указано, берется last_feature.
      - save_as (str, опционально): имя для сохранения операции фаски.
    """
    part = runner.ctx["part"]
    distance = params.get("distance", 1.0)
    use_feature = params.get("use_feature")
    save_as = params.get("save_as")
    
    feature_to_use = None
    if use_feature and "saved_features" in runner.ctx:
        feature_to_use = runner.ctx["saved_features"].get(use_feature)
    if not feature_to_use:
        feature_to_use = runner.ctx.get("last_feature")
        
    if not feature_to_use:
        print("Нет базового элемента для фаски.")
        return

    import win32com.client
    try:
        feature7 = win32com.client.CastTo(feature_to_use, 'IFeature7')
    except Exception:
        feature7 = feature_to_use

    try:
        edges = feature7.GetModelObjects(7) # 7 = o3d_edge
    except AttributeError:
        edges = feature7.ModelObjects(7)

    if edges:
        model_container = runner.module7.IModelContainer(part)
        chamfers = model_container.Chamfers
        chamfer = chamfers.Add()
        
        edges_list = list(edges)
        chamfer.BaseObjects = edges_list
        chamfer.Distance1 = distance
        chamfer.Distance2 = distance
        chamfer.Update()
        
        runner.ctx["last_feature"] = chamfer
        if save_as:
            if "saved_features" not in runner.ctx:
                runner.ctx["saved_features"] = {}
            runner.ctx["saved_features"][save_as] = chamfer
            
        print(f"Построена фаска {distance} мм на {len(edges_list)} ребрах" + (f" (сохранено как '{save_as}')" if save_as else ""))
    else:
        print("Не найдено ребер для фаски")


def op_fillet_all_edges(runner, params):
    """
    Строит скругление на всех ребрах указанной операции.
    
    Параметры (params):
      - radius (float, по умолчанию 5.0): радиус скругления.
      - use_feature (str, опционально): имя операции, ребра которой нужно скруглить.
      - save_as (str, опционально): имя для сохранения операции скругления.
    """
    part = runner.ctx["part"]
    radius = params.get("radius", 5.0)
    use_feature = params.get("use_feature")
    save_as = params.get("save_as")
    
    feature_to_use = None
    if use_feature and "saved_features" in runner.ctx:
        feature_to_use = runner.ctx["saved_features"].get(use_feature)
    if not feature_to_use:
        feature_to_use = runner.ctx.get("last_feature")
        
    if not feature_to_use:
        print("Нет базового элемента для скругления.")
        return

    import win32com.client
    try:
        feature7 = win32com.client.CastTo(feature_to_use, 'IFeature7')
    except Exception:
        feature7 = feature_to_use

    try:
        edges = feature7.GetModelObjects(7)
    except AttributeError:
        edges = feature7.ModelObjects(7)

    if edges:
        model_container = runner.module7.IModelContainer(part)
        fillets = model_container.Fillets
        fillet = fillets.Add()
        fillet.Radius1 = radius
        fillet.BaseObjects = list(edges)
        fillet.Update()
        
        runner.ctx["last_feature"] = fillet
        if save_as:
            if "saved_features" not in runner.ctx:
                runner.ctx["saved_features"] = {}
            runner.ctx["saved_features"][save_as] = fillet
            
        print(f"Выполнено скругление {len(list(edges))} ребер радиусом {radius} мм" + (f" (сохранено как '{save_as}')" if save_as else ""))
    else:
        print("Не найдено ребер для скругления")


def op_add_thread(runner, params):
    """
    Добавляет условное обозначение резьбы на цилиндрическую грань указанного элемента.
    
    Параметры (params):
      - length (float, по умолчанию 10.0): длина резьбы.
      - pitch (float, по умолчанию 1.0): шаг резьбы.
      - use_feature (str, опционально): имя цилиндрического элемента (например стержня), на который наносится резьба.
      - save_as (str, опционально): имя для сохранения операции резьбы.
    """
    part = runner.ctx["part"]
    length = params.get("length", 10.0)
    pitch = params.get("pitch", 1.0)
    use_feature = params.get("use_feature")
    save_as = params.get("save_as")
    
    feature_to_use = None
    if use_feature and "saved_features" in runner.ctx:
        feature_to_use = runner.ctx["saved_features"].get(use_feature)
    if not feature_to_use:
        feature_to_use = runner.ctx.get("last_feature")
        
    if not feature_to_use:
        print("Нет базового элемента для нанесения резьбы.")
        return

    import win32com.client
    import pythoncom
    
    try:
        feature7 = win32com.client.CastTo(feature_to_use, 'IFeature7')
    except Exception:
        feature7 = feature_to_use

    try:
        faces = feature7.GetModelObjects(6)
        edges = feature7.GetModelObjects(5)
    except AttributeError:
        faces = feature7.ModelObjects(6)
        edges = feature7.ModelObjects(5)

    if not faces:
        print("Не найдены грани для нанесения резьбы.")
        return

    symbols_container = part._oleobj_.QueryInterface(
        runner.module7.NamesToIIDMap['ISymbols3DContainer'], 
        pythoncom.IID_IDispatch
    )
    symbols_container = runner.module7.ISymbols3DContainer(symbols_container)
    
    threads = symbols_container.Threads
    thread = threads.Add()
    thread.AutoLenght = False
    thread.Lenght = length
    thread.AutoDiameter = True 
    
    try:
        thread_params = win32com.client.CastTo(thread, 'IThreadsParameters')
        if thread_params:
            thread_params.Pitch = pitch
    except Exception:
        pass

    success = False
    for face in faces:
        thread.BaseObject = face
        if thread.Update():
            success = True
            break
        if edges:
            for edge in edges:
                thread.InitialBorder = edge
                if thread.Update():
                    success = True
                    break
        if success:
            break

    if success:
        runner.ctx["last_feature"] = thread
        if save_as:
            if "saved_features" not in runner.ctx:
                runner.ctx["saved_features"] = {}
            runner.ctx["saved_features"][save_as] = thread
        part.Update() 
        print(f"Построено изображение резьбы (длина {length} мм, шаг {pitch} мм)" + (f" (сохранено как '{save_as}')" if save_as else ""))
    else:
        print("Ошибка: среди граней элемента не найдена цилиндрическая поверхность.")


def op_create_sketch_on_face(runner, params):
    """
    Создает эскиз на указанной плоской грани элемента.
    
    Параметры (params):
      - face_index (int/str, по умолчанию "top"): Индекс грани или "top" (самая высокая по оси Z), "-1" (последняя).
      - use_feature (str, опционально): Имя операции (выдавливания), у которой ищется грань. ОЧЕНЬ ВАЖНО УКАЗЫВАТЬ!
      - save_as (str, опционально): Имя для сохранения нового эскиза.
    """
    part = runner.ctx["part"]
    face_index_param = params.get("face_index", "top") 
    use_feature = params.get("use_feature")
    save_as = params.get("save_as")
    
    feature_to_use = None
    if use_feature and "saved_features" in runner.ctx:
        feature_to_use = runner.ctx["saved_features"].get(use_feature)
        if not feature_to_use:
            print(f"Предупреждение: сохраненная операция '{use_feature}' не найдена. Будет использована последняя.")
            
    if not feature_to_use:
        feature_to_use = runner.ctx.get("last_feature")
        
    if not feature_to_use:
        print("Нет операции для выбора грани.")
        return

    import win32com.client
    try:
        feature7 = win32com.client.CastTo(feature_to_use, 'IFeature7')
    except Exception:
        feature7 = feature_to_use

    try:
        faces = feature7.GetModelObjects(6)
    except AttributeError:
        faces = feature7.ModelObjects(6)

    if not faces:
        print("Грани не найдены.")
        return

    base_face = None
    faces_count = len(faces)

    if str(face_index_param).lower() == "top":
        best_face = None
        max_z = -1e9
        found_by_z = False

        for face in faces:
            try:
                imo = win32com.client.CastTo(face, 'IModelObject')
                edges = imo.GetModelObjects(5) if hasattr(imo, 'GetModelObjects') else imo.ModelObjects(5)
                
                if edges and len(edges) > 0:
                    edge = win32com.client.CastTo(edges[0], 'IEdge')
                    try:
                        vertex = edge.Vertex(True)
                        if vertex:
                            res, x, y, z = vertex.GetPoint(0.0, 0.0, 0.0)
                            if res and z > max_z:
                                max_z = z
                                best_face = face
                                found_by_z = True
                    except:
                        pass
            except Exception:
                continue

        if found_by_z and best_face:
            base_face = best_face
            print(f"Автоматически найдена верхняя грань (Z = {max_z:.2f})")
        else:
            print(f"Не удалось извлечь координаты, беру последнюю грань: индекс [{faces_count - 1}].")
            base_face = faces[faces_count - 1]
            
    else:
        idx = int(face_index_param)
        if idx < 0:
            idx = faces_count + idx
        if 0 <= idx < faces_count:
            base_face = faces[idx]
        else:
            base_face = faces[faces_count - 1]

    model_container = runner.module7.IModelContainer(part)
    sketches = model_container.Sketchs
    sketch = sketches.Add()
    sketch.Plane = base_face
    sketch.Update()
    sketch.BeginEdit()
    
    runner.ctx["sketch"] = sketch
    if save_as:
        if "saved_features" not in runner.ctx:
            runner.ctx["saved_features"] = {}
        runner.ctx["saved_features"][save_as] = sketch
        
    print(f"Создан эскиз на грани" + (f" (сохранен как '{save_as}')" if save_as else ""))


def op_set_material_and_color(runner, params):
    """
    Устанавливает материал и цвет детали.
    
    Параметры (params):
      - material (str, по умолчанию "Сталь 10 ГОСТ 1050-88"): название материала.
      - density (float, по умолчанию 7.85): плотность.
      - color (list[int], опционально): цвет RGB в формате [R, G, B].
    """
    part = runner.ctx["part"]
    material_name = params.get("material", "Сталь 10 ГОСТ 1050-88")
    density = params.get("density", 7.85)
    
    # Установка материала
    part.SetMaterial(material_name, density)
    
    # Установка цвета (RGB)
    if "color" in params:
        import win32com.client
        try:
            # Получаем интерфейс управления цветом детали
            color_param = win32com.client.CastTo(part, 'IColorParam7')
            r, g, b = params["color"]
            
            # В КОМПАСе цвет задается как одно число (0xBBGGRR)
            kompas_color = (b << 16) | (g << 8) | r
            
            # Присваиваем цвет
            color_param.Color = kompas_color
            
            # 0 = useColorOur (использовать собственный заданный цвет, а не цвет материала)
            color_param.UseColor = 0 
            
        except Exception as e:
            print(f"Не удалось установить цвет: {e}")
    
    part.Update()
    print(f"Установлен материал: {material_name}" + (" и обновлен цвет" if "color" in params else ""))



def op_save_document(runner, params):
    """
    Сохраняет 3D документ на диск.
    
    Параметры (params):
      - path (str, по умолчанию "cube.m3d"): путь для сохранения. Если относительный, сохранится рядом со скриптом.
    """
    import os
    path = params.get("path", "cube.m3d")
    
    if not os.path.isabs(path):
        path = os.path.abspath(path)
        
    doc = runner.app.ActiveDocument
    doc.SaveAs(path)
    print(f"Документ сохранен по пути: {path}")

# def op_draw_line(runner, params):
#     """
#     Рисует отрезок (линию) в текущем эскизе.
#     Если не указана точка start, линия начинается с конца предыдущей линии.
#     Конечную точку можно задать абсолютно (end) или относительно (dx, dy).
    
#     Параметры:
#       - start (list[float], опционально): [X, Y] начала. Если нет, берется last_point.
#       - end (list[float], опционально): [X, Y] конца.
#       - dx (float), dy (float) (опционально): смещение относительно начала (вместо end).
#     """
#     # Определяем начальную точку
#     if "start" in params:
#         x1, y1 = params["start"]
#     elif "last_point" in runner.ctx:
#         x1, y1 = runner.ctx["last_point"]
#     else:
#         x1, y1 = 0.0, 0.0

#     # Сохраняем первую точку контура, если её еще нет
#     if "first_point" not in runner.ctx:
#         runner.ctx["first_point"] = [x1, y1]

#     # Определяем конечную точку
#     if "end" in params:
#         x2, y2 = params["end"]
#     elif "dx" in params or "dy" in params:
#         dx = params.get("dx", 0.0)
#         dy = params.get("dy", 0.0)
#         x2 = x1 + dx
#         y2 = y1 + dy
#     else:
#         raise ValueError("Для draw_line необходимо указать 'end' [X, Y] или 'dx'/'dy'")

#     import win32com.client
#     app7 = win32com.client.Dispatch("Kompas.Application.7")
#     sketch_doc = runner.ctx.get("sketch_doc", app7.ActiveDocument)
#     doc2d = win32com.client.CastTo(sketch_doc, 'IKompasDocument2D')
#     view = doc2d.ViewsAndLayersManager.Views.ActiveView
#     drawing_container = win32com.client.CastTo(view, 'IDrawingContainer')

#     # Рисуем линию
#     lines = drawing_container.LineSegments
#     line = lines.Add()
#     line.X1 = x1
#     line.Y1 = y1
#     line.X2 = x2
#     line.Y2 = y2
#     line.Update()

#     # Обновляем последнюю точку
#     runner.ctx["last_point"] = [x2, y2]
    
#     print(f"Нарисован отрезок от ({x1}, {y1}) до ({x2}, {y2})")


# def op_close_contour(runner, params):
#     """
#     Замыкает контур: рисует линию от последней точки (last_point) к первой (first_point).
#     """
#     if "last_point" not in runner.ctx or "first_point" not in runner.ctx:
#         print("Невозможно замкнуть контур: нет начальных или конечных точек.")
#         return
        
#     x1, y1 = runner.ctx["last_point"]
#     x2, y2 = runner.ctx["first_point"]
    
#     # Чтобы не дублировать код API, можно просто вызвать op_draw_line
#     # Но мы нарисуем напрямую для надежности:
#     import win32com.client
#     app7 = win32com.client.Dispatch("Kompas.Application.7")
#     sketch_doc = runner.ctx.get("sketch_doc", app7.ActiveDocument)
#     doc2d = win32com.client.CastTo(sketch_doc, 'IKompasDocument2D')
#     drawing_container = win32com.client.CastTo(doc2d.ViewsAndLayersManager.Views.ActiveView, 'IDrawingContainer')
    
#     lines = drawing_container.LineSegments
#     line = lines.Add()
#     line.X1 = x1
#     line.Y1 = y1
#     line.X2 = x2
#     line.Y2 = y2
#     line.Update()
    
#     runner.ctx["last_point"] = [x2, y2]
#     print(f"Контур замкнут линией от ({x1}, {y1}) до ({x2}, {y2})")
    
def op_rib(runner, params):
    """
    Создает ребро жесткости по текущему эскизу (обычно это эскиз с одной линией).
    
    Параметры (params):
      - thickness (float, по умолчанию 5.0): толщина ребра.
      - side (int, по умолчанию 2): направление утолщения 
        (0 - в одну сторону, 1 - в другую, 2 - симметрично относительно эскиза).
      - reverse_direction (bool, по умолчанию False): сменить направление построения ребра (в сторону тела).
      - save_as (str, опционально): имя для сохранения элемента.
    """
    part = runner.ctx.get("part")
    sketch = runner.ctx.get("sketch")
    
    thickness = params.get("thickness", 5.0)
    side = params.get("side", 2)
    reverse_dir = params.get("reverse_direction", False)
    save_as = params.get("save_as")
    
    if not part or not sketch:
        print("Ошибка: Нет активной детали или эскиза для ребра жесткости.")
        return
        
    model_container = runner.module7.IModelContainer(part)
    ribs = model_container.Ribs
    rib = ribs.Add()
    
    # 1. Привязываем эскиз
    rib.Sketch = sketch
    
    import win32com.client
    # 2. Настраиваем параметры тонкой стенки (толщина и симметрия)
    try:
        thin_params = win32com.client.CastTo(rib, 'IThinParameters')
        thin_params.Thin = True
        
        # 2 - симметрично, 0 и 1 - в разные стороны
        if side == 2:
            thin_params.Thickness1 = thickness / 2.0
            thin_params.Thickness2 = thickness / 2.0
        elif side == 0:
            thin_params.Thickness1 = thickness
            thin_params.Thickness2 = 0.0
        else:
            thin_params.Thickness1 = 0.0
            thin_params.Thickness2 = thickness
    except Exception as e:
        print(f"Предупреждение при настройке толщины ребра: {e}")

    # 3. Направление "дотягивания" до тела
    # В разных версиях API это может быть Direction, ReverseDirection или свойство внутри IExtrusion
    try:
        if hasattr(rib, 'Direction'):
            rib.Direction = 1 if reverse_dir else 0
        elif hasattr(rib, 'ReverseDirection'):
            rib.ReverseDirection = reverse_dir
    except:
        pass

    rib.Update()
    
    runner.ctx["last_feature"] = rib
    if save_as:
        if "saved_features" not in runner.ctx:
            runner.ctx["saved_features"] = {}
        runner.ctx["saved_features"][save_as] = rib
        
    print(f"Построено ребро жесткости (толщина {thickness} мм)")
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
def op_create_sketch_on_plane(runner, params):
    """
    Создает эскиз на базовой или смещенной плоскости и открывает его для редактирования.
    """
    part = runner.ctx.get("part")
    if not part:
        print("Ошибка: Нет активной детали.")
        return

    plane_name = params.get("plane", "XY")
    offset = float(params.get("offset", 0.0)) # Убеждаемся, что это число

    import win32com.client
    
    # 1. Получаем базовую плоскость
    try:
        const = win32com.client.gencache.EnsureModule("{75C9F5D0-B5BF-11D3-BF26-005084D136BD}", 0, 1, 0).constants
        plane_map = {"XY": const.o3d_planeXOY, "ZX": const.o3d_planeXOZ, "ZY": const.o3d_planeYOZ}
        plane_offset_type = const.o3d_planeOffset
    except:
        plane_map = {"XY": 1, "ZX": 2, "ZY": 3}
        plane_offset_type = 18

    plane_type = plane_map.get(plane_name, plane_map["XY"])
    base_plane = part.DefaultObject(plane_type)

    target_plane = base_plane
    
    # 2. Если есть смещение, создаем смещенную плоскость
    if offset != 0.0:
        try:
            aux_geom = win32com.client.CastTo(part, 'IAuxiliaryGeomContainer')
            if aux_geom:
                planes = aux_geom.Planes3D
                plane_obj = planes.Add(plane_offset_type) 
                
                plane_params = win32com.client.CastTo(plane_obj, 'IPlane3DByOffset')
                
                # ЯВНОЕ ПРИВЕДЕНИЕ базовой плоскости к интерфейсу IModelObject
                base_plane_model = win32com.client.CastTo(base_plane, 'IModelObject')
                plane_params.BasePlane = base_plane_model
                
                # Явно передаем значения нужных типов
                plane_params.Offset = abs(offset)
                plane_params.Direction = True if offset > 0 else False
                
                plane_obj.Update()
                target_plane = plane_obj
                print(f"Создана смещенная плоскость от {plane_name} (смещение {offset} мм)")
            else:
                print("Не удалось получить IAuxiliaryGeomContainer")
        except Exception as e:
            print(f"Предупреждение: Не удалось создать смещенную плоскость: {e}")
            target_plane = base_plane

    # 3. Создаем эскиз
    sketches = runner.module7.IModelContainer(part).Sketchs
    sketch = sketches.Add()
    
    # Также явно приводим целевую плоскость к IModelObject для эскиза (Type Mismatch proof)
    target_plane_model = win32com.client.CastTo(target_plane, 'IModelObject')
    sketch.Plane = target_plane_model
    sketch.Update()
    
    runner.ctx["sketch"] = sketch
    
    # 4. Входим в режим редактирования
    try:
        sketch_doc = sketch.BeginEdit()
        runner.ctx["sketch_doc"] = sketch_doc
        runner.ctx["current_contour_points"] = []
        if offset == 0.0:
            print(f"Создан эскиз на базовой плоскости {plane_name}")
        else:
            print(f"Создан эскиз на смещенной плоскости")
    except Exception as e:
        print(f"Ошибка при входе в эскиз: {e}")
        
        
        
        
        
        
def op_draw_line(runner, params):
    """
    Рисует линию в текущем ОТКРЫТОМ эскизе.
    """
    sketch_doc = runner.ctx.get("sketch_doc")
    if not sketch_doc:
        print("Ошибка: Нет открытого эскиза (sketch_doc).")
        return
        
    start = params.get("start")
    end = params.get("end")
    dx = params.get("dx")
    dy = params.get("dy")
    
    points = runner.ctx.get("current_contour_points", [])
    
    
    if start is None:
        if points:
            start = points[-1]
        else:
            start = [0.0, 0.0]
            
    if end is None:
        if dx is not None and dy is not None:
            end = [start[0] + dx, start[1] + dy]
        else:
            print("Ошибка: недостаточно параметров для draw_line")
            return
            
    import win32com.client
    # В API7 документ эскиза (IFragmentDocument) поддерживает интерфейс IKompasDocument2D
    doc2d = win32com.client.CastTo(sketch_doc, 'IKompasDocument2D')
    if doc2d:
        try:
            # Получаем менеджер видов 2D документа эскиза
            views = doc2d.ViewsAndLayersManager.Views
            active_view = views.ActiveView
            drawing_container = win32com.client.CastTo(active_view, 'IDrawingContainer')
            lines = drawing_container.LineSegments
            line = lines.Add()
            line.X1, line.Y1 = start[0], start[1]
            line.X2, line.Y2 = end[0], end[1]
            # line.Style = params.get("style", 1) 
            line.Update()
            
            # Сохраняем точки для продолжения контура
            if not points:
                points.append(start)
            points.append(end)
            runner.ctx["current_contour_points"] = points
            print(f"Нарисован отрезок от {start} до {end}")
        except Exception as e:
            print(f"Ошибка при рисовании линии: {e}")
    else:
        print("Не удалось получить IKompasDocument2D из эскиза")

def op_close_contour(runner, params):
    """
    Замыкает контур текущего эскиза (соединяет последнюю точку с первой).
    """
    points = runner.ctx.get("current_contour_points", [])
    if len(points) < 3:
        print("Недостаточно точек для замыкания контура.")
        return
        
    start = points[-1]
    end = points[0]
    
    # Просто вызываем op_draw_line с нужными координатами
    op_draw_line(runner, {"start": start, "end": end})
    print("Контур замкнут")

def op_finish_sketch(runner, params):
    """
    Завершает редактирование эскиза (EndEdit) и может сохранить его в контекст.
    """
    sketch = runner.ctx.get("sketch")
    if sketch:
        try:
            sketch.EndEdit()
            runner.ctx.pop("sketch_doc", None)
            runner.ctx.pop("current_contour_points", None)
            
            # Сохраняем эскиз под именем, если передано save_as
            save_as = params.get("save_as")
            if save_as:
                runner.ctx[save_as] = sketch
                print(f"Эскиз завершен и сохранен под именем '{save_as}'.")
            else:
                print("Эскиз завершен (EndEdit).")
        except Exception as e:
            print(f"Ошибка при завершении эскиза: {e}")
            sketch.Update()
            
def _get_axis_object(part, axis_name):
    """
    Вспомогательная функция: создает объект 3D-оси (OX/OY/OZ) во вспомогательной геометрии
    и возвращает его как IModelObject, пригодный для использования в IRotated.Axis.
    """
    import win32com.client

    # Получаем константы КОМПАС
    try:
        const = win32com.client.gencache.EnsureModule(
            "{75C9F5D0-B5BF-11D3-BF26-005084D136BD}", 0, 1, 0
        ).constants
        axis_type_map = {
            "X": const.o3d_axisOX,
            "Y": const.o3d_axisOY,
            "Z": const.o3d_axisOZ,
        }
    except Exception:
        # fallback: числовые коды осей (на случай проблем с gencache)
        axis_type_map = {
            "X": 4,  # o3d_axisOX
            "Y": 5,  # o3d_axisOY
            "Z": 6,  # o3d_axisOZ
        }

    aux_geom = win32com.client.CastTo(part, "IAuxiliaryGeomContainer")
    axes3d = aux_geom.Axes3D
    axis_obj = axes3d.Add(axis_type_map[axis_name])
    axis_obj.Update()

    # Явно приводим к IModelObject
    try:
        axis_model = win32com.client.CastTo(axis_obj, "IModelObject")
    except Exception:
        axis_model = axis_obj

    return axis_model


def op_revolve_sketch(runner, params):
    """
    Вращение текущего эскиза вокруг оси, заданной отрезком в этом же эскизе.
    Берется первый найденный отрезок как ось вращения.
    
    Параметры:
      - angle (float, по умолчанию 360.0): угол вращения в градусах.
      - save_as (str, опционально): имя для сохранения операции.
    """
    part = runner.ctx.get("part")
    sketch = runner.ctx.get("sketch")

    if not part or not sketch:
        print("Ошибка: нет активной детали или эскиза для операции вращения.")
        return

    import win32com.client

    angle = float(params.get("angle", 360.0))
    save_as = params.get("save_as")

    # 1. Получаем документ эскиза и ищем в нем первый отрезок (ось вращения)
    sketch_doc = runner.ctx.get("sketch_doc")
    if not sketch_doc:
        print("Ошибка: нет открытого документа эскиза (sketch_doc) для поиска оси.")
        return

    try:
        doc2d = win32com.client.CastTo(sketch_doc, "IKompasDocument2D")
        views = doc2d.ViewsAndLayersManager.Views
        active_view = views.ActiveView
        drawing_container = win32com.client.CastTo(active_view, "IDrawingContainer")
        lines = drawing_container.LineSegments
        if lines.Count == 0:
            print("Ошибка: в эскизе нет отрезков для использования в качестве оси вращения.")
            return
        axis_line = lines.Item(0)
    except Exception as e:
        print(f"Ошибка при поиске отрезка-оси в эскизе: {e}")
        return

    # 2. Создаем элемент вращения в API7
    model_container = runner.module7.IModelContainer(part)
    rotateds = model_container.Rotateds

    # 28 = o3d_bossRotated (элемент вращения)
    rotated = rotateds.Add(28)
    if rotated is None:
        print("Ошибка: не удалось создать операцию вращения.")
        return

    try:
        i_rotated = win32com.client.CastTo(rotated, "IRotated")
    except Exception:
        i_rotated = rotated

    # 3. Назначаем профиль (эскиз)
    sketch_assigned = False
    try:
        i_rotated.Profile = sketch
        sketch_assigned = True
    except Exception:
        pass

    if not sketch_assigned:
        try:
            i_rotated.Sketch = sketch
            sketch_assigned = True
        except Exception:
            pass

    if not sketch_assigned:
        print("Ошибка: не удалось назначить эскиз в операцию вращения.")
        return

    # 4. Назначаем угол
    if abs(angle - 360.0) < 0.1:
        try:
            i_rotated.Torus = True
        except Exception:
            pass
    else:
        try:
            i_rotated.Angle1 = angle
        except Exception:
            pass

    # 5. Назначаем ось вращения как направляющий объект (отрезок из эскиза)
    try:
        i_rotated.Axis = axis_line
    except Exception as e:
        print(f"Ошибка при назначении оси вращения по отрезку: {e}")
        return

    # 6. Строим операцию
    if not rotated.Update():
        print("Ошибка: операция вращения не построена. Проверьте замкнутость эскиза и положение оси.")
        return

    runner.ctx["last_feature"] = rotated
    if save_as:
        if "saved_features" not in runner.ctx:
            runner.ctx["saved_features"] = {}
        runner.ctx["saved_features"][save_as] = rotated

    print(f"Выполнено вращение эскиза на {angle}° вокруг оси, заданной отрезком эскиза"
          + (f" (сохранено как '{save_as}')" if save_as else ""))