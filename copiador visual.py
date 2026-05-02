from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton,
    QLabel, QFileDialog, QCheckBox,
    QProgressBar, QTextEdit, QListWidget,
    QListWidgetItem, QMessageBox, QLineEdit,
    QGroupBox, QScrollArea, QWidget, QVBoxLayout,
    QDialog, QHBoxLayout, QFormLayout, QGridLayout,
)
from PyQt5.QtGui import QIcon, QCursor
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMimeData, QSettings
import os
import sys
import shutil
import zipfile
import tempfile
import openpyxl
import time
import json
import re
from datetime import datetime

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

class ResizableMessageDialog(QDialog):
    def __init__(self, title, message, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(500, 300)
        self.setModal(True)
        
        layout = QVBoxLayout()
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setText(message)
        scroll_area.setWidget(text_edit)
        layout.addWidget(scroll_area)

        self.yes_button = QPushButton("Sí")
        self.no_button = QPushButton("No")
        layout.addWidget(self.yes_button)
        layout.addWidget(self.no_button)

        self.setLayout(layout)
        self.yes_button.clicked.connect(self.accept)
        self.no_button.clicked.connect(self.reject)
        self.setSizeGripEnabled(True)

class WorkerCopia(QThread):
    progreso_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finalizado_signal = pyqtSignal()
    error_signal = pyqtSignal(str)
    archivo_actual_signal = pyqtSignal(str)
    notificacion_signal = pyqtSignal(str)
    pregunta_signal = pyqtSignal(str, int)
    correccion_signal = pyqtSignal(str, str, dict)
    pregunta_directorios_signal = pyqtSignal(str, list)

    def __init__(self, directorios_origen, directorio_destino, copiar_todo, opciones_copia, texto_reemplazo_rips, texto_sufijo_rips, texto_reemplazo_docker, texto_sufijo_docker, copiar_en_raiz, comprimir_zip, aut_compensar, texto_reemplazo_zip, texto_sufijo_zip, texto_reemplazo_directorios, texto_sufijo_directorios):
        super().__init__()
        self.directorios_origen = directorios_origen
        self.directorio_destino = directorio_destino
        self.copiar_todo = copiar_todo
        self.opciones_copia = opciones_copia
        self.texto_reemplazo_rips = texto_reemplazo_rips
        self.texto_sufijo_rips = texto_sufijo_rips
        self.texto_reemplazo_docker = texto_reemplazo_docker
        self.texto_sufijo_docker = texto_sufijo_docker
        self.copiar_en_raiz = copiar_en_raiz
        self.comprimir_zip = comprimir_zip
        self.aut_compensar = aut_compensar
        self.texto_reemplazo_zip = texto_reemplazo_zip
        self.texto_sufijo_zip = texto_sufijo_zip
        self.texto_reemplazo_directorios = texto_reemplazo_directorios
        self.texto_sufijo_directorios = texto_sufijo_directorios
        self.formato_json = opciones_copia.get('formato_json', False)
        self.cancelar_flag = False
        self.archivos_copiados = 0
        self.total_archivos = 0
        self.directorios_copiados = {}
        self.archivos_corregidos_copiados = set()
        self.todos_los_errores = []
        self.validacion_autorizaciones = {}
        self.respuesta_pregunta = {}
        self.archivos_encontrados = []  # <---- AÑADIDA
        self.directorios_a_copiar = set() # AÑADIDA

    def find_all_keys_with_context(self, obj, target, context=None, results=None, path=None):
        if results is None:
            results = []
        if path is None:
            path = []
        if context is None:
            context = "raíz"
        if isinstance(obj, dict):
            for key, value in obj.items():
                new_context = f"{context}.{key}" if context != "raíz" else key
                new_path = path + [key]
                if key.lower() == target.lower():
                    identifier = None
                    if "codConsulta" in obj:
                        identifier = f"\"codConsulta\": \"{obj['codConsulta']}\""
                    elif "nomTecnologiaSalud" in obj:
                        identifier = f"\"nomTecnologiaSalud\": \"{obj['nomTecnologiaSalud']}\""
                    elif "codProcedimiento" in obj:
                        identifier = f"\"codProcedimiento\": \"{obj['codProcedimiento']}\""
                    elif "codTecnologiaSalud" in obj:
                        identifier = f"\"codTecnologiaSalud\": \"{obj['codTecnologiaSalud']}\""
                    results.append((value, new_context.split('.')[1], identifier, new_path))
                self.find_all_keys_with_context(value, target, new_context, results, new_path)
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                new_context = f"{context}[{i}]"
                new_path = path + [i]
                self.find_all_keys_with_context(item, target, new_context, results, new_path)
        return results

    def run(self):
        try:
            self.tiempo_inicio = time.time()
            self.log_signal.emit("Iniciando proceso de copia de directorios...")
            self.log_signal.emit(f"Directorios origen iniciales: {self.directorios_origen}")

            self.archivos_encontrados = []
            self.buscar_archivos()  # Llamada a la nueva función de búsqueda
            if self.cancelar_flag:
                self.log_signal.emit("Proceso cancelado por el usuario durante la búsqueda.")
                self.finalizado_signal.emit()
                return

            self.directorios_a_validar = list(self.directorios_origen) # Inicializamos con todos los directorios origen
            self.directorios_a_copiar = set() # Se llenará en validar_archivos
            self.todos_los_errores = []
            self.validacion_autorizaciones = {}
            self.respuesta_pregunta = {}

            if not self.copiar_todo:
                self.validar_archivos()  # Llamada a la nueva función de validación
                if self.cancelar_flag:
                    self.log_signal.emit("Proceso cancelado por el usuario durante la validación.")
                    self.finalizado_signal.emit()
                    return
                self.log_signal.emit(f"Directorios aprobados para copiar tras validaciones: {self.directorios_a_copiar}")
            else:
                self.log_signal.emit("Copiando todos los directorios sin validaciones posteriores a la búsqueda inicial.")
                self.directorios_a_copiar = set(self.directorios_origen)

            self.preparar_conteo_archivos(self.directorios_a_copiar)

            # Renombrado y copia de directorios
            directorios_con_renombrado = {}
            for dir_origen in self.directorios_a_copiar:
                dir_name = os.path.basename(dir_origen)
                nuevo_nombre_dir = dir_name
                if self.opciones_copia['renombrar_directorios']:
                    prefijo = self.texto_reemplazo_directorios if self.texto_reemplazo_directorios else ""
                    sufijo = self.texto_sufijo_directorios if self.texto_sufijo_directorios else ""
                    nuevo_nombre_dir = f"{prefijo}{dir_name}{sufijo}"
                    self.log_signal.emit(f"Renombrando directorio: {dir_name} -> {nuevo_nombre_dir}")
                if self.copiar_en_raiz:
                    dir_destino = self.directorio_destino
                else:
                    dir_destino = os.path.join(self.directorio_destino, nuevo_nombre_dir)
                directorios_con_renombrado[dir_origen] = dir_destino

            self.archivos_copiados = 0
            for dir_origen, dir_destino in directorios_con_renombrado.items():
                if self.cancelar_flag:
                    break
                if self.comprimir_zip:
                    temp_dir = tempfile.mkdtemp()
                    archivos_copiados_en_zip = self.copiar_archivos(dir_origen, temp_dir, self.copiar_todo, self.opciones_copia, self.copiar_en_raiz)
                    if archivos_copiados_en_zip > 0:
                        zip_name = os.path.basename(dir_origen)
                        if self.opciones_copia['renombrar_zip']:
                            prefijo = self.texto_reemplazo_zip if self.texto_reemplazo_zip else ""
                            sufijo = self.texto_sufijo_zip if self.texto_sufijo_zip else ""
                            zip_name = f"{prefijo}{zip_name}{sufijo}"
                            self.log_signal.emit(f"Renombrando ZIP: {os.path.basename(dir_origen)} -> {zip_name}")
                        zip_path = os.path.join(self.directorio_destino, f"{zip_name}.zip")
                        shutil.make_archive(zip_path[:-4], 'zip', temp_dir)
                    shutil.rmtree(temp_dir)
                else:
                    archivos_copiados_en_directorio = self.copiar_archivos(dir_origen, dir_destino, self.copiar_todo, self.opciones_copia, self.copiar_en_raiz)
                    if not hasattr(self, 'directorios_copiados'):
                        self.directorios_copiados = {}
                    self.directorios_copiados[dir_origen] = archivos_copiados_en_directorio
                    self.archivos_copiados += archivos_copiados_en_directorio

                if self.aut_compensar and os.path.basename(dir_origen) in self.validacion_autorizaciones:
                    self.todos_los_errores.extend(self.validacion_autorizaciones[os.path.basename(dir_origen)])

                progreso = int((self.archivos_copiados / self.total_archivos) * 100) if self.total_archivos > 0 else 0
                self.progreso_signal.emit(progreso)

            if self.todos_los_errores and not self.copiar_todo:
                mensaje_descarga = "Se encontraron los siguientes problemas:\n" + "\n".join(self.todos_los_errores) + "\n\n¿Desea descargar el archivo con todos los errores?"
                self.respuesta_usuario = None
                self.pregunta_signal.emit(mensaje_descarga, QMessageBox.Yes | QMessageBox.No)
                while self.respuesta_usuario is None and not self.cancelar_flag:
                    time.sleep(0.1)
                if self.respuesta_usuario == QMessageBox.Yes:
                    archivo_validacion = os.path.join(self.directorio_destino, "Validacion_Problemas.txt")
                    with open(archivo_validacion, 'w', encoding='utf-8') as f:
                        f.write("Problemas de Validación Detectados:\n")
                        for error in self.todos_los_errores:
                            f.write(f"{error}\n")
                    self.log_signal.emit(f"Archivo de validación guardado en {archivo_validacion}")

            self.log_signal.emit(f"Copia finalizada en {time.time() - self.tiempo_inicio:.2f} segundos.")
            self.finalizado_signal.emit()
        except Exception as e:
            self.error_signal.emit(f"Error durante la ejecución principal: {str(e)}")
            self.finalizado_signal.emit()

    def buscar_archivos(self):
        self.mensaje.emit("Buscando archivos en los directorios de origen...")
        for directorio_origen in self.directorios_origen:
            if self.detener_copia:
                break
            self.mensaje.emit(f"Buscando en: {directorio_origen}")
            try:
                for nombre_archivo in os.listdir(directorio_origen):
                    if nombre_archivo.startswith("ResultadosDoker_") and nombre_archivo.endswith(".json"):
                        self.archivos_encontrados.append(os.path.join(directorio_origen, nombre_archivo))
            except Exception as e:
                self.mensaje.emit(f"Error al acceder a {directorio_origen}: {e}")
        self.mensaje.emit(f"Se encontraron {len(self.archivos_encontrados)} archivos para validar.")

    def validar_archivos(self):
        
        self.mensaje.emit("Validando directorios y archivos...")
        directorios_rips_invalidos = set()
        directorios_docker_invalidos = set()
        directorios_otros_invalidos = set()
        directorios_a_copiar_temp = set()

        directorios_procesados = set() # Para evitar procesar el mismo directorio varias veces

        for archivo_path in self.archivos_encontrados:
            if self.detener_copia:
                break
            directorio_origen = os.path.dirname(archivo_path)
            if directorio_origen in directorios_procesados:
                continue
            directorios_procesados.add(directorio_origen)

            valido = True

            # Validación de RIPS
            rips_path = os.path.join(directorio_origen, "RIPS")
            if not os.path.isdir(rips_path):
                directorios_rips_invalidos.add(directorio_origen)
                valido = False

            # Validación de Docker
            try:
                with open(archivo_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if data.get("ResultState") != "OK":
                        directorios_docker_invalidos.add(directorio_origen)
                        valido = False
            except Exception as e:
                self.mensaje.emit(f"Error al leer o validar {archivo_path}: {e}")
                directorios_docker_invalidos.add(directorio_origen)
                valido = False

            # Aquí podrías añadir otras validaciones para el directorio 'directorio_origen'
            # ...

            if valido:
                directorios_a_copiar_temp.add(directorio_origen)

        # Preguntar al usuario sobre los directorios problemáticos (después de la validación)
        directorios_a_considerar = set(self.directorios_origen) # Considerar todos los directorios origen

        if directorios_rips_invalidos:
            mensaje_rips = f"Se encontraron directorios sin la carpeta 'RIPS':\n{chr(10).join(directorios_rips_invalidos)}\n¿Desea intentar copiar estos directorios?"
            dialogo_rips = ResizableMessageDialog("Advertencia: Carpetas sin RIPS", mensaje_rips, parent=self.parent())
            if dialogo_rips.exec_() == QDialog.Accepted:
                directorios_a_considerar.update(directorios_rips_invalidos)

        if directorios_docker_invalidos:
            mensaje_docker = f"Se encontraron directorios con archivos Docker no 'OK':\n{chr(10).join(directorios_docker_invalidos)}\n¿Desea intentar copiar estos directorios?"
            dialogo_docker = ResizableMessageDialog("Advertencia: Resultados Docker no OK", mensaje_docker, parent=self.parent())
            if dialogo_docker.exec_() == QDialog.Accepted:
                directorios_a_considerar.update(directorios_docker_invalidos)

        # Finalmente, actualiza la lista de directorios a copiar basándote en los directorios originales
        # que pasaron las validaciones o fueron aprobados por el usuario.
        self.directorios_a_copiar = directorios_a_considerar.intersection(directorios_a_copiar_temp.union(directorios_rips_invalidos, directorios_docker_invalidos))

        self.mensaje.emit(f"Se encontraron {len(self.directorios_a_copiar)} directorios para copiar.")
    def calcular_ruta_destino(self, ruta_origen, dir_origen, nombre_directorio):
        file = os.path.basename(ruta_origen)
        ruta_relativa = os.path.relpath(ruta_origen, dir_origen)
        nombre_archivo_original, extension_archivo = os.path.splitext(file)

        if self.copiar_todo:
            return os.path.join(self.directorio_destino, nombre_directorio, ruta_relativa)
        elif self.copiar_en_raiz:
            if file.startswith("Rips_") and self.texto_reemplazo_rips:
                texto_reemplazo = self.texto_reemplazo_rips if self.texto_reemplazo_rips else "Rips"
                texto_sufijo = self.texto_sufijo_rips if self.texto_sufijo_rips else ""
                filename_parts = [texto_reemplazo, nombre_archivo_original[len('Rips_'):]]
                if texto_sufijo:
                    filename_parts.append(texto_sufijo)
                nuevo_nombre = "_".join(filter(None, filename_parts)) + extension_archivo
                return os.path.join(self.directorio_destino, nuevo_nombre)
            return os.path.join(self.directorio_destino, file)
        else:
            if file.startswith("Rips_") and self.texto_reemplazo_rips:
                texto_reemplazo = self.texto_reemplazo_rips if self.texto_reemplazo_rips else "Rips"
                texto_sufijo = self.texto_sufijo_rips if self.texto_sufijo_rips else ""
                filename_parts = [texto_reemplazo, nombre_archivo_original[len('Rips_'):]]
                if texto_sufijo:
                    filename_parts.append(texto_sufijo)
                nuevo_nombre = "_".join(filter(None, filename_parts)) + extension_archivo
                return os.path.join(self.directorio_destino, nombre_directorio, nuevo_nombre)
            return os.path.join(self.directorio_destino, nombre_directorio, file)

    def preparar_conteo_archivos(self):
        total_count = 0
        for dir_origen in self.directorios_origen:
            for root, _, files in os.walk(dir_origen):
                for file in files:
                    ruta_origen = os.path.join(root, file)
                    if self.debe_copiar_archivo(ruta_origen, root, file, self.copiar_todo, self.opciones_copia, os.path.basename(dir_origen), f"Rips_{os.path.basename(dir_origen)}.json")[0]:
                        total_count += 1
        self.total_archivos = total_count
        self.archivos_copiados = 0

    def copiar_archivos(self, dir_origen, dir_destino, copiar_todo, opciones_copia, copiar_en_raiz):
        nombre_base = os.path.basename(dir_destino)
        archivo_rips = f"Rips_{os.path.basename(dir_origen)}.json"
        archivos_copiados = []
        self.validacion_autorizaciones = getattr(self, 'validacion_autorizaciones', {})

        self.log_signal.emit(f"Procesando directorio: {nombre_base}")
        self.log_signal.emit(f"Directorio destino: {dir_destino}")

        self.total_archivos = sum(len(files) for _, _, files in os.walk(dir_origen))

        for root, _, files in os.walk(dir_origen):
            for file in files:
                if self.cancelar_flag:
                    self.log_signal.emit("Copia cancelada en copiar_archivos.")
                    return archivos_copiados
                ruta_origen = os.path.join(root, file)
                debe_copiar, tipo_archivo = self.debe_copiar_archivo(ruta_origen, root, file, copiar_todo, opciones_copia, os.path.basename(dir_origen), archivo_rips)
                if debe_copiar:
                    es_rips = file.startswith("Rips_")
                    es_docker = file.startswith("ResultadosDoker_")
                    txt_rips = opciones_copia.get('txt_rips', False)
                    txt_docker = opciones_copia.get('txt_docker', False)

                    nombre_archivo_final = file
                    nombre_directorio = os.path.basename(dir_origen)
                    
                    if es_rips and opciones_copia.get('renombrar_rips', False):
                        prefijo = self.texto_reemplazo_rips if self.texto_reemplazo_rips else ""
                        sufijo = self.texto_sufijo_rips if self.texto_sufijo_rips else ""
                        ext = '.txt' if txt_rips else os.path.splitext(file)[1]
                        nombre_archivo_final = f"{prefijo}{nombre_directorio}{sufijo}{ext}"
                        self.log_signal.emit(f"DEBUG RENOMBRADO RIPS: prefijo='{prefijo}', sufijo='{sufijo}', base='{nombre_directorio}', ext='{ext}'")
                    
                    elif es_docker and opciones_copia.get('renombrar_docker', False):
                        prefijo = self.texto_reemplazo_docker if self.texto_reemplazo_docker else ""
                        sufijo = self.texto_sufijo_docker if self.texto_sufijo_docker else ""
                        ext = '.txt' if txt_docker else os.path.splitext(file)[1]
                        nombre_archivo_final = f"{prefijo}{nombre_directorio}{sufijo}{ext}"
                        self.log_signal.emit(f"DEBUG RENOMBRADO DOCKER: prefijo='{prefijo}', sufijo='{sufijo}', base='{nombre_directorio}', ext='{ext}'")
                    
                    elif (es_rips and txt_rips) or (es_docker and txt_docker):
                        nombre_archivo_final = os.path.splitext(file)[0] + '.txt'
                        self.log_signal.emit(f"Cambiando extensión a .txt: {file} -> {nombre_archivo_final}")
                    else:
                        self.log_signal.emit(f"Renombrado no aplicado a {file}: condiciones no cumplidas")

                    if copiar_en_raiz:
                        ruta_destino_final = os.path.join(dir_destino, nombre_archivo_final)
                        self.log_signal.emit(f"**RUTA DESTINO (RAIZ ACTIVADA):** {ruta_destino_final}")
                    else:
                        ruta_destino_final = os.path.join(dir_destino, nombre_archivo_final)
                        self.log_signal.emit(f"**RUTA DESTINO (RAIZ DESACTIVADA):** {ruta_destino_final}")

                    if not os.path.exists(os.path.dirname(ruta_destino_final)):
                        os.makedirs(os.path.dirname(ruta_destino_final), exist_ok=True)
                        self.log_signal.emit(f"Creado directorio: {os.path.dirname(ruta_destino_final)}")

                    es_json = os.path.splitext(file)[1].lower() == '.json' and (es_rips or es_docker)

                    try:
                        if ruta_origen not in self.archivos_corregidos_copiados:
                            self.archivo_actual_signal.emit(file)

                            data = None
                            if tipo_archivo == "docker_file" and not self.copiar_todo:
                                for encoding in ['utf-8', 'latin-1', 'windows-1252']:
                                    try:
                                        with open(ruta_origen, 'r', encoding=encoding) as f:
                                            data = json.load(f)
                                        break
                                    except UnicodeDecodeError:
                                        continue
                                    except json.JSONDecodeError as e:
                                        self.log_signal.emit(f"Omitiendo {file}: Error al leer JSON ({str(e)})")
                                        break
                                if data is None:
                                    self.log_signal.emit(f"Omitiendo {file}: No se pudo decodificar o leer como JSON con ninguna codificación soportada")
                                    continue

                                result_state_key = next((key for key in data.keys() if key.lower() == "resultstate"), None)
                                if not result_state_key or not data[result_state_key]:
                                    self.log_signal.emit(f"Omitiendo {file}: ResultState no es True o no encontrado")
                                    continue

                            if self.aut_compensar and tipo_archivo == "rips_json":
                                if data is None:
                                    for encoding in ['utf-8', 'latin-1', 'windows-1252']:
                                        try:
                                            with open(ruta_origen, 'r', encoding=encoding) as f:
                                                data = json.load(f)
                                            break
                                        except UnicodeDecodeError:
                                            continue
                                        except json.JSONDecodeError as e:
                                            self.log_signal.emit(f"Omitiendo {file}: Error al leer JSON ({str(e)})")
                                            break
                                    if data is None:
                                        self.log_signal.emit(f"Omitiendo {file}: No se pudo decodificar o leer como JSON con ninguna codificación soportada")
                                        continue

                                servicios = data.get('usuarios', [{}])[0].get('servicios', {})
                                correcciones = {}
                                problemas_autorizaciones = []
                                servicio_types = ['consultas', 'medicamentos', 'procedimientos', 'otrosServicios', 'hospitalizacion']
                                for servicio_type in servicio_types:
                                    if servicio_type in servicios:
                                        for i, servicio in enumerate(servicios[servicio_type]):
                                            num_autorizacion = servicio.get('numAutorizacion')
                                            if num_autorizacion is not None and num_autorizacion != "":
                                                if not (num_autorizacion.isdigit() and len(num_autorizacion) == 15):
                                                    digits = re.sub(r'[^0-9]', '', str(num_autorizacion))
                                                    corrected = digits[:15] if len(digits) > 15 else digits.ljust(15, '0')
                                                    path = (servicio_type, i, 'numAutorizacion')
                                                    correcciones[path] = corrected
                                                    problemas_autorizaciones.append(f"{servicio_type}.{i}.numAutorizacion: Original='{num_autorizacion}', Propuesto='{corrected}'")

                                if correcciones:
                                    self.respuesta_correccion = None
                                    self.correccion_signal.emit(file, ruta_origen, correcciones)
                                    while self.respuesta_correccion is None and not self.cancelar_flag:
                                        time.sleep(0.1)
                                    if self.respuesta_correccion == QMessageBox.Yes:
                                        for path, corrected_value in correcciones.items():
                                            servicio_type, idx, key = path
                                            servicios[servicio_type][idx][key] = corrected_value
                                        self.log_signal.emit(f"Correcciones aplicadas en {file}: {correcciones}")
                                        problemas_autorizaciones.append("Corrección aceptada")
                                    else:
                                        self.log_signal.emit(f"Correcciones rechazadas para {file}")
                                        problemas_autorizaciones.append("Corrección rechazada")
                                    self.validacion_autorizaciones[os.path.basename(dir_origen)] = problemas_autorizaciones

                            if es_json:
                                if data is None:
                                    for encoding in ['utf-8', 'latin-1', 'windows-1252']:
                                        try:
                                            with open(ruta_origen, 'r', encoding=encoding) as f:
                                                data = json.load(f)
                                            break
                                        except UnicodeDecodeError:
                                            continue
                                        except json.JSONDecodeError as e:
                                            self.log_signal.emit(f"Omitiendo {file}: Error al leer JSON ({str(e)})")
                                            break
                                    if data is None:
                                        self.log_signal.emit(f"Omitiendo {file}: No se pudo decodificar o leer como JSON con ninguna codificación soportada")
                                        continue

                                if (es_rips and txt_rips) or (es_docker and txt_docker):
                                    with open(ruta_destino_final, 'w', encoding='utf-8') as f:
                                        f.write(json.dumps(data, indent=4, ensure_ascii=False))
                                    self.log_signal.emit(f"Copiado como TXT: {ruta_origen} a {ruta_destino_final}")
                                elif self.formato_json:
                                    with open(ruta_destino_final, 'w', encoding='utf-8') as f:
                                        json.dump(data, f, indent=4, ensure_ascii=False)
                                    self.log_signal.emit(f"Copiado y formateado como JSON: {ruta_origen} a {ruta_destino_final}")
                                else:
                                    with open(ruta_destino_final, 'w', encoding='utf-8') as f:
                                        json.dump(data, f, ensure_ascii=False)
                                    self.log_signal.emit(f"Copiado como JSON: {ruta_origen} a {ruta_destino_final}")
                            else:
                                shutil.copy2(ruta_origen, ruta_destino_final)
                                self.log_signal.emit(f"Copiado: {ruta_origen} a {ruta_destino_final}")

                            if os.path.exists(ruta_destino_final):
                                self.log_signal.emit(f"Archivo confirmado en disco: {ruta_destino_final}")
                            else:
                                self.log_signal.emit(f"¡ERROR! Archivo no encontrado en disco después de copiar: {ruta_destino_final}")
                            dir_contenido = os.listdir(os.path.dirname(ruta_destino_final))
                            self.log_signal.emit(f"Contenido de {os.path.dirname(ruta_destino_final)} después de copiar {file}: {dir_contenido}")

                            self.archivos_copiados += 1
                            progreso = int((self.archivos_copiados / max(self.total_archivos, 1)) * 100)
                            self.progreso_signal.emit(progreso)
                            archivos_copiados.append(ruta_destino_final)
                    except Exception as e:
                        error_msg = f"Error al copiar {ruta_origen} a {ruta_destino_final}: {str(e)}"
                        self.error_signal.emit(error_msg)
                        self.log_signal.emit(error_msg)
                else:
                    self.log_signal.emit(f"Omitido por filtro: {ruta_origen}")

        self.log_signal.emit(f"Archivos copiados: {archivos_copiados}")
        return archivos_copiados

    def debe_copiar_archivo(self, ruta_origen, root, file, copiar_todo=False, opciones_copia=None, nombre_directorio_param=None, archivo_rips_param=None):
        if copiar_todo:
            self.log_signal.emit(f"DEBUG (COPY CHECK): Copiando todo - {file}: True")
            return True, None
        
        if opciones_copia is None:
            self.log_signal.emit(f"DEBUG (COPY CHECK): Opciones nulas - {file}: False")
            return False, None
        
        if not any(opciones_copia.values()):
            self.log_signal.emit(f"DEBUG (COPY CHECK): Ninguna opción activa - {file}: True")
            return True, None

        nombre_directorio = nombre_directorio_param
        archivo_rips = archivo_rips_param
        debe_copiar = False
        tipo_archivo = None

        if opciones_copia['rips_archivo'] and file == archivo_rips:
            debe_copiar = True
            tipo_archivo = "rips_json"
            self.log_signal.emit(f"DEBUG (RIPS CHECK): {file} == {archivo_rips} - Resultado: True (rips_json)")
        elif opciones_copia['rips_archivo']:
            self.log_signal.emit(f"DEBUG (RIPS CHECK): {file} != {archivo_rips} - Resultado: False")

        if opciones_copia['rips_docker'] and file.startswith("ResultadosDoker_") and file.endswith(".json"):
            debe_copiar = True
            tipo_archivo = "docker_file"
            self.log_signal.emit(f"DEBUG (DOCKER CHECK): {file} es ResultadosDoker_*.json - Resultado: True (docker_file)")

        if opciones_copia['ad_xml'] and file.startswith("ad") and file.endswith(".xml"):
            debe_copiar = True
            tipo_archivo = None
            self.log_signal.emit(f"DEBUG (AD CHECK): {file} es ad*.xml - Resultado: True")

        if opciones_copia['ar_xml'] and file.startswith("ar") and file.endswith(".xml"):
            debe_copiar = True
            tipo_archivo = None
            self.log_signal.emit(f"DEBUG (AR CHECK): {file} es ar*.xml - Resultado: True")

        if opciones_copia['fv_pdf'] and file.startswith("fv") and file.endswith(".pdf"):
            debe_copiar = True
            tipo_archivo = None
            self.log_signal.emit(f"DEBUG (FV PDF CHECK): {file} es fv*.pdf - Resultado: True")

        if opciones_copia['fv_xml'] and file.startswith("fv") and file.endswith(".xml"):
            debe_copiar = True
            tipo_archivo = None
            self.log_signal.emit(f"DEBUG (FV XML CHECK): {file} es fv*.xml - Resultado: True")

        if not debe_copiar:
            self.log_signal.emit(f"DEBUG (COPY CHECK): {file} no cumple ningún criterio - Resultado: False")

        return debe_copiar, tipo_archivo

class CopiadorDirectorios(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Copiador de Directorios")
        self.setMinimumSize(950, 950)  # Tamaño mínimo para la ventana
        self.setWindowIcon(QIcon(resource_path("iconos_2/kurama.png")))
        self.settings = QSettings("MiEmpresa", "CopiadorDirectorios")

        # Widget central y layout principal
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Sección superior: Origen, Destino y Opciones adicionales
        top_section_layout = QHBoxLayout()
        main_layout.addLayout(top_section_layout)

        # Columna izquierda: Origen y Destino (más pequeña)
        origen_destino_layout = QVBoxLayout()
        
        self.origen_input = QLineEdit("")
        self.origen_input.textChanged.connect(self.actualizar_origen)
        self.boton_origen = QPushButton("Sel.")  # Botón más pequeño
        self.boton_origen.setFixedWidth(50)  # Reducir ancho del botón
        self.boton_origen.setIcon(QIcon(resource_path("iconos_2/carpeta.png")))
        self.boton_origen.clicked.connect(self.seleccionar_origen)
        origen_row = QHBoxLayout()
        origen_row.addWidget(QLabel("Origen:"))
        origen_row.addWidget(self.origen_input)
        origen_row.addWidget(self.boton_origen)
        origen_destino_layout.addLayout(origen_row)

        self.destino_input = QLineEdit("")
        self.destino_input.textChanged.connect(self.actualizar_destino)
        self.boton_destino = QPushButton("Sel.")  # Botón más pequeño
        self.boton_destino.setFixedWidth(50)  # Reducir ancho del botón
        self.boton_destino.setIcon(QIcon(resource_path("iconos_2/carpeta.png")))
        self.boton_destino.clicked.connect(self.seleccionar_destino)
        destino_row = QHBoxLayout()
        destino_row.addWidget(QLabel("Destino:"))
        destino_row.addWidget(self.destino_input)
        destino_row.addWidget(self.boton_destino)
        origen_destino_layout.addLayout(destino_row)

        # Ajustar tamaño de la columna izquierda
        top_section_layout.addLayout(origen_destino_layout)
        top_section_layout.setStretch(0, 1)  # Columna izquierda ocupa 1/3 del espacio

        # Columna derecha: Opciones adicionales (más grande y reorganizada)
        extra_options_group = QGroupBox("Opciones Adicionales")
        extra_options_layout = QGridLayout()  # Cambiar a QGridLayout para mejor organización
        extra_options_group.setLayout(extra_options_layout)
        
        self.check_2025 = QCheckBox("Restringir 2025")
        self.check_txt_rips = QCheckBox("txt Rips")
        self.check_txt_docker = QCheckBox("txt Docker")
        self.check_tema_oscuro = QCheckBox("Tema Claro")
        self.check_activar_Copiarenraiz = QCheckBox("Copiar en raíz")
        self.check_formato_json = QCheckBox("Formato JSON")
        self.check_comprimir_zip = QCheckBox("Comprimir ZIP")
        self.check_comprimir_zip.stateChanged.connect(self.actualizar_estado_comprimir)
        self.check_aut_compensar = QCheckBox("Aut Compensar")
        self.check_copiar_sin_validar = QCheckBox("Sin validar")

        # Reorganizar checkboxes en una cuadrícula 3x3
        extra_options_layout.addWidget(self.check_2025, 0, 0)
        extra_options_layout.addWidget(self.check_txt_rips, 0, 1)
        extra_options_layout.addWidget(self.check_txt_docker, 0, 2)
        extra_options_layout.addWidget(self.check_tema_oscuro, 1, 0)
        extra_options_layout.addWidget(self.check_activar_Copiarenraiz, 1, 1)
        extra_options_layout.addWidget(self.check_formato_json, 1, 2)
        extra_options_layout.addWidget(self.check_comprimir_zip, 2, 0)
        extra_options_layout.addWidget(self.check_aut_compensar, 2, 1)
        extra_options_layout.addWidget(self.check_copiar_sin_validar, 2, 2)

        top_section_layout.addWidget(extra_options_group)
        top_section_layout.setStretch(1, 2)  # Columna derecha ocupa 2/3 del espacio

        # Sección de opciones de copia (checkboxes de archivos)
        file_options_group = QGroupBox("Opciones de Copia")
        file_options_layout = QGridLayout()
        file_options_group.setLayout(file_options_layout)
        main_layout.addWidget(file_options_group)

        self.check_copiar_ad_xml = QCheckBox("ad.xml")
        self.check_copiar_ar_xml = QCheckBox("ar.xml")
        self.check_copiar_fv_pdf = QCheckBox("fv.pdf")
        self.check_copiar_fv_xml = QCheckBox("fv.xml")
        self.check_copiar_rips_archivo = QCheckBox("RIPS")
        self.check_copiar_rips_docker = QCheckBox("DOKER")

        file_options_layout.addWidget(self.check_copiar_ad_xml, 0, 0)
        file_options_layout.addWidget(self.check_copiar_ar_xml, 0, 1)
        file_options_layout.addWidget(self.check_copiar_fv_pdf, 0, 2)
        file_options_layout.addWidget(self.check_copiar_fv_xml, 0, 3)
        file_options_layout.addWidget(self.check_copiar_rips_archivo, 1, 0)
        file_options_layout.addWidget(self.check_copiar_rips_docker, 1, 1)

        # Sección de renombrado
        renaming_section_layout = QHBoxLayout()
        main_layout.addLayout(renaming_section_layout)

        # Renombrado RIPS
        self.grupo_renombrado_rips = QGroupBox("Renombrado RIPS")
        rips_layout = QFormLayout()
        self.grupo_renombrado_rips.setLayout(rips_layout)
        self.check_activar_renombrado_rips = QCheckBox("Renombrar RIPS")
        self.check_activar_renombrado_rips.stateChanged.connect(self.actualizar_estado_renombrado_rips)
        self.input_reemplazo_rips = QLineEdit()
        self.input_reemplazo_rips.setPlaceholderText("Opcional: Reemplazar 'Rips'")
        self.input_sufijo_rips = QLineEdit()
        self.input_sufijo_rips.setPlaceholderText("Opcional: Añadir sufijo")
        rips_layout.addRow(self.check_activar_renombrado_rips)
        rips_layout.addRow(QLabel("Prefijo:"), self.input_reemplazo_rips)
        rips_layout.addRow(QLabel("Sufijo:"), self.input_sufijo_rips)
        renaming_section_layout.addWidget(self.grupo_renombrado_rips)

        # Renombrado Docker
        self.grupo_renombrado_docker = QGroupBox("Renombrado Docker")
        docker_layout = QFormLayout()
        self.grupo_renombrado_docker.setLayout(docker_layout)
        self.check_activar_renombrado_docker = QCheckBox("Renombrar Docker")
        self.check_activar_renombrado_docker.stateChanged.connect(self.actualizar_estado_renombrado_docker)
        self.input_reemplazo_docker = QLineEdit()
        self.input_reemplazo_docker.setPlaceholderText("Opcional: Reemplazar 'ResultadosDoker'")
        self.input_sufijo_docker = QLineEdit()
        self.input_sufijo_docker.setPlaceholderText("Opcional: Añadir sufijo")
        docker_layout.addRow(self.check_activar_renombrado_docker)
        docker_layout.addRow(QLabel("Prefijo:"), self.input_reemplazo_docker)
        docker_layout.addRow(QLabel("Sufijo:"), self.input_sufijo_docker)
        renaming_section_layout.addWidget(self.grupo_renombrado_docker)

        # Renombrado ZIP
        self.grupo_renombrado_zip = QGroupBox("Renombrado ZIP")
        zip_layout = QFormLayout()
        self.grupo_renombrado_zip.setLayout(zip_layout)
        self.check_activar_renombrado_zip = QCheckBox("Renombrar ZIP")
        self.check_activar_renombrado_zip.stateChanged.connect(self.actualizar_estado_renombrado_zip)
        self.input_reemplazo_zip = QLineEdit()
        self.input_reemplazo_zip.setPlaceholderText("Opcional: Reemplazar en nombre del ZIP")
        self.input_sufijo_zip = QLineEdit()
        self.input_sufijo_zip.setPlaceholderText("Opcional: Añadir sufijo al ZIP")
        zip_layout.addRow(self.check_activar_renombrado_zip)
        zip_layout.addRow(QLabel("Prefijo:"), self.input_reemplazo_zip)
        zip_layout.addRow(QLabel("Sufijo:"), self.input_sufijo_zip)
        renaming_section_layout.addWidget(self.grupo_renombrado_zip)

        # Renombrado Directorios
        self.grupo_renombrado_directorios = QGroupBox("Renombrado Directorios")
        dirs_layout = QFormLayout()
        self.grupo_renombrado_directorios.setLayout(dirs_layout)
        self.check_activar_renombrado_directorios = QCheckBox("Renombrar Directorios")
        self.check_activar_renombrado_directorios.stateChanged.connect(self.actualizar_estado_renombrado_directorios)
        self.input_reemplazo_directorios = QLineEdit()
        self.input_reemplazo_directorios.setPlaceholderText("Opcional: Añadir prefijo al directorio")
        self.input_sufijo_directorios = QLineEdit()
        self.input_sufijo_directorios.setPlaceholderText("Opcional: Añadir sufijo al directorio")
        dirs_layout.addRow(self.check_activar_renombrado_directorios)
        dirs_layout.addRow(QLabel("Prefijo:"), self.input_reemplazo_directorios)
        dirs_layout.addRow(QLabel("Sufijo:"), self.input_sufijo_directorios)
        renaming_section_layout.addWidget(self.grupo_renombrado_directorios)

        # Sección de lista y entrada manual
        list_manual_section_layout = QHBoxLayout()
        main_layout.addLayout(list_manual_section_layout)

        # Lista de directorios
        list_layout = QVBoxLayout()
        self.lista_directorios_copiar = QListWidget()
        self.lista_directorios_copiar.setSelectionMode(QListWidget.MultiSelection)
        self.lista_directorios_copiar.setAcceptDrops(True)
        list_layout.addWidget(self.lista_directorios_copiar)

        list_buttons_layout = QHBoxLayout()
        self.boton_agregar_listado = QPushButton("Agregar Listado (TXT/Excel)")
        self.boton_agregar_listado.setIcon(QIcon(resource_path("iconos_2/nota.png")))
        self.boton_agregar_listado.clicked.connect(self.agregar_listado)
        self.boton_eliminar_listado = QPushButton("Eliminar Listado")
        self.boton_eliminar_listado.setIcon(QIcon(resource_path("iconos_2/eliminar.png")))
        self.boton_eliminar_listado.clicked.connect(self.eliminar_listado)
        list_buttons_layout.addWidget(self.boton_agregar_listado)
        list_buttons_layout.addWidget(self.boton_eliminar_listado)
        list_layout.addLayout(list_buttons_layout)
        list_manual_section_layout.addLayout(list_layout)

        # Entrada manual
        manual_layout = QVBoxLayout()
        self.label_entrada_manual = QLabel("Entrada Manual de Directorios:")
        self.manual_input_text = QTextEdit()
        self.manual_input_text.setPlaceholderText("Ingrese los directorios manualmente, uno por línea")
        self.manual_input_text.textChanged.connect(self.actualizar_contador_manual)
        self.label_contador_manual = QLabel("Directorios: 0")
        self.label_contador_manual.setAlignment(Qt.AlignRight | Qt.AlignBottom)
        self.boton_agregar_manual = QPushButton("Agregar Manualmente")
        self.boton_agregar_manual.setIcon(QIcon(resource_path("iconos_2/editar.png")))
        self.boton_agregar_manual.clicked.connect(self.agregar_listado_manual)
        manual_layout.addWidget(self.label_entrada_manual)
        manual_layout.addWidget(self.manual_input_text)
        manual_layout.addWidget(self.label_contador_manual)
        manual_layout.addWidget(self.boton_agregar_manual)
        list_manual_section_layout.addLayout(manual_layout)

        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Progreso: 0%")
        main_layout.addWidget(self.progress_bar)

        # Etiqueta de archivo actual
        self.label_archivo_actual = QLabel("Archivo actual: ")
        self.label_archivo_actual.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.label_archivo_actual)

        # Área de logs
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

        # Botones de acción
        buttons_layout = QHBoxLayout()
        self.boton_copiar = QPushButton("Copiar Directorios")
        self.boton_copiar.setIcon(QIcon(resource_path("iconos_2/copiar.png")))
        self.boton_copiar.clicked.connect(self.confirmar_copia)
        self.boton_copiar.setEnabled(False)
        self.boton_cancelar = QPushButton("Cancelar")
        self.boton_cancelar.setIcon(QIcon(resource_path("iconos_2/cancelar.png")))
        self.boton_cancelar.clicked.connect(self.cancelar_copia)
        self.boton_cancelar.setEnabled(False)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.boton_copiar)
        buttons_layout.addWidget(self.boton_cancelar)
        buttons_layout.addStretch()
        main_layout.addLayout(buttons_layout)

        # Información del creador
        self.label_informacion = QLabel("Creado por: Armando Villa G- contacto: 3158171343")
        self.label_informacion.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.label_informacion)

        # Estirar el layout para ocupar el espacio disponible
        main_layout.addStretch()

        # Conexiones y configuraciones iniciales
        self.check_tema_oscuro.stateChanged.connect(self.cambiar_tema)
        self.setAcceptDrops(True)
        self.cargar_configuraciones()
        self.check_tema_oscuro.setChecked(self.settings.value("tema_oscuro", False, type=bool))
        self.cambiar_tema()
        self.worker_thread = None
        self.inicio_copia = None
        self.opciones_copia = {}
        self.actualizar_estado_renombrado_rips()
        self.actualizar_estado_renombrado_docker()
        self.actualizar_estado_renombrado_zip()
        self.actualizar_estado_renombrado_directorios()
        self.actualizar_estado_comprimir()

    def showEvent(self, event):
        self.resize(950, 950)
        super().showEvent(event)

    def actualizar_estado_comprimir(self):
        zip_enabled = self.check_comprimir_zip.isChecked() and self.check_activar_renombrado_zip.isChecked()
        self.input_reemplazo_zip.setEnabled(zip_enabled)
        self.input_sufijo_zip.setEnabled(zip_enabled)
        self.check_activar_renombrado_zip.setEnabled(self.check_comprimir_zip.isChecked())

    def cargar_configuraciones(self):
        ruta_origen = self.settings.value("ruta_origen", "", type=str)
        if ruta_origen.startswith("Directorio de Origen: "):
            ruta_origen = ruta_origen.replace("Directorio de Origen: ", "")
        self.origen_input.setText(ruta_origen)

        ruta_destino = self.settings.value("ruta_destino", "", type=str)
        if ruta_destino.startswith("Directorio de Destino: "):
            ruta_destino = ruta_destino.replace("Directorio de Destino: ", "")
        self.destino_input.setText(ruta_destino)

        self.check_copiar_ad_xml.setChecked(self.settings.value("check_ad_xml", False, type=bool))
        self.check_copiar_ar_xml.setChecked(self.settings.value("check_ar_xml", False, type=bool))
        self.check_copiar_fv_pdf.setChecked(self.settings.value("check_fv_pdf", False, type=bool))
        self.check_copiar_fv_xml.setChecked(self.settings.value("check_fv_xml", False, type=bool))
        self.check_copiar_rips_docker.setChecked(self.settings.value("check_rips_docker", False, type=bool))
        self.check_copiar_rips_archivo.setChecked(self.settings.value("check_rips_archivo", False, type=bool))
        self.check_activar_renombrado_rips.setChecked(self.settings.value("check_renombrado_rips", False, type=bool))
        self.check_activar_renombrado_docker.setChecked(self.settings.value("check_renombrado_docker", False, type=bool))
        self.check_activar_renombrado_zip.setChecked(self.settings.value("check_renombrado_zip", False, type=bool))
        self.check_activar_renombrado_directorios.setChecked(self.settings.value("check_renombrado_directorios", False, type=bool))
        self.check_activar_Copiarenraiz.setChecked(self.settings.value("check_copiar_en_raiz", False, type=bool))
        self.check_comprimir_zip.setChecked(self.settings.value("check_comprimir_zip", False, type=bool))
        self.check_aut_compensar.setChecked(self.settings.value("check_aut_compensar", False, type=bool))
        self.check_copiar_sin_validar.setChecked(self.settings.value("check_copiar_sin_validar", False, type=bool))
        self.input_reemplazo_rips.setText(self.settings.value("reemplazo_rips", "", type=str))
        self.input_sufijo_rips.setText(self.settings.value("sufijo_rips", "", type=str))
        self.input_reemplazo_docker.setText(self.settings.value("reemplazo_docker", "", type=str))
        self.input_sufijo_docker.setText(self.settings.value("sufijo_docker", "", type=str))
        self.input_reemplazo_zip.setText(self.settings.value("reemplazo_zip", "", type=str))
        self.input_sufijo_zip.setText(self.settings.value("sufijo_zip", "", type=str))
        self.input_reemplazo_directorios.setText(self.settings.value("reemplazo_directorios", "", type=str))
        self.input_sufijo_directorios.setText(self.settings.value("sufijo_directorios", "", type=str))

        self.boton_copiar.setEnabled(
            self.origen_input.text().strip() != "" and 
            self.destino_input.text().strip() != "" and 
            self.lista_directorios_copiar.count() > 0
        )

    def toggle_renombrado_fields(self, checked):
        rips_enabled = checked and self.check_activar_renombrado_rips.isChecked()
        self.input_reemplazo_rips.setEnabled(rips_enabled)
        self.input_sufijo_rips.setEnabled(rips_enabled)
        self.check_activar_renombrado_rips.setEnabled(checked)

        docker_enabled = checked and self.check_activar_renombrado_docker.isChecked()
        self.input_reemplazo_docker.setEnabled(docker_enabled)
        self.input_sufijo_docker.setEnabled(docker_enabled)
        self.check_activar_renombrado_docker.setEnabled(checked)

        zip_enabled = checked and self.check_comprimir_zip.isChecked() and self.check_activar_renombrado_zip.isChecked()
        self.input_reemplazo_zip.setEnabled(zip_enabled)
        self.input_sufijo_zip.setEnabled(zip_enabled)
        self.check_activar_renombrado_zip.setEnabled(checked and self.check_comprimir_zip.isChecked())

        dirs_enabled = checked and self.check_activar_renombrado_directorios.isChecked()
        self.input_reemplazo_directorios.setEnabled(dirs_enabled)
        self.input_sufijo_directorios.setEnabled(dirs_enabled)
        self.check_activar_renombrado_directorios.setEnabled(checked)

    def closeEvent(self, event):
        self.settings.setValue("ruta_origen", self.origen_input.text())
        self.settings.setValue("ruta_destino", self.destino_input.text())
        self.settings.setValue("check_ad_xml", self.check_copiar_ad_xml.isChecked())
        self.settings.setValue("check_ar_xml", self.check_copiar_ar_xml.isChecked())
        self.settings.setValue("check_fv_pdf", self.check_copiar_fv_pdf.isChecked())
        self.settings.setValue("check_fv_xml", self.check_copiar_fv_xml.isChecked())
        self.settings.setValue("check_rips_docker", self.check_copiar_rips_docker.isChecked())
        self.settings.setValue("check_rips_archivo", self.check_copiar_rips_archivo.isChecked())
        self.settings.setValue("check_renombrado_rips", self.check_activar_renombrado_rips.isChecked())
        self.settings.setValue("reemplazo_rips", self.input_reemplazo_rips.text())
        self.settings.setValue("sufijo_rips", self.input_sufijo_rips.text())
        self.settings.setValue("check_renombrado_docker", self.check_activar_renombrado_docker.isChecked())
        self.settings.setValue("reemplazo_docker", self.input_reemplazo_docker.text())
        self.settings.setValue("sufijo_docker", self.input_sufijo_docker.text())
        self.settings.setValue("check_renombrado_zip", self.check_activar_renombrado_zip.isChecked())
        self.settings.setValue("reemplazo_zip", self.input_reemplazo_zip.text())
        self.settings.setValue("sufijo_zip", self.input_sufijo_zip.text())
        self.settings.setValue("check_renombrado_directorios", self.check_activar_renombrado_directorios.isChecked())
        self.settings.setValue("reemplazo_directorios", self.input_reemplazo_directorios.text())
        self.settings.setValue("sufijo_directorios", self.input_sufijo_directorios.text())
        self.settings.setValue("check_copiar_en_raiz", self.check_activar_Copiarenraiz.isChecked())
        self.settings.setValue("check_comprimir_zip", self.check_comprimir_zip.isChecked())
        self.settings.setValue("check_aut_compensar", self.check_aut_compensar.isChecked())
        self.settings.setValue("check_copiar_sin_validar", self.check_copiar_sin_validar.isChecked())
        self.settings.setValue("tema_oscuro", self.check_tema_oscuro.isChecked())
        event.accept()

    def confirmar_copia(self):
        print("Iniciando nueva ejecución de copia...")

        directorios_seleccionados = []
        print(f"Total de ítems en lista_directorios_copiar: {self.lista_directorios_copiar.count()}")
        for i in range(self.lista_directorios_copiar.count()):
            item = self.lista_directorios_copiar.item(i)
            print(f"Ítem {i}: {item.text()} (Seleccionado: {item.isSelected()})")
            if item.isSelected():
                # Usar la ruta completa almacenada en UserRole
                ruta_completa = item.data(Qt.UserRole)
                if ruta_completa and os.path.isdir(ruta_completa):
                    directorios_seleccionados.append(ruta_completa)
                    print(f"Directorio seleccionado: {ruta_completa}")
                else:
                    print(f"Advertencia: Ruta inválida para {item.text()}: {ruta_completa}")

        if not directorios_seleccionados:
            QMessageBox.warning(self, "Advertencia", "No hay directorios seleccionados válidos para copiar.")
            print("No hay directorios seleccionados válidos.")
            return

        directorio_destino = self.destino_input.text().strip()
        if not directorio_destino:
            QMessageBox.warning(self, "Advertencia", "Primero seleccione el directorio de destino.")
            print("Directorio de destino no seleccionado.")
            return

        opciones_copia = {
            'ad_xml': self.check_copiar_ad_xml.isChecked(),
            'ar_xml': self.check_copiar_ar_xml.isChecked(),
            'fv_pdf': self.check_copiar_fv_pdf.isChecked(),
            'fv_xml': self.check_copiar_fv_xml.isChecked(),
            'rips_docker': self.check_copiar_rips_docker.isChecked(),
            'rips_archivo': self.check_copiar_rips_archivo.isChecked(),
            'copiar_en_raiz': self.check_activar_Copiarenraiz.isChecked(),
            'comprimir_zip': self.check_comprimir_zip.isChecked(),
            'renombrar_rips': self.check_activar_renombrado_rips.isChecked(),
            'reemplazo_rips': self.input_reemplazo_rips.text().strip(),
            'sufijo_rips': self.input_sufijo_rips.text().strip(),
            'renombrar_docker': self.check_activar_renombrado_docker.isChecked(),
            'reemplazo_docker': self.input_reemplazo_docker.text().strip(),
            'sufijo_docker': self.input_sufijo_docker.text().strip(),
            'renombrar_zip': self.check_activar_renombrado_zip.isChecked(),
            'reemplazo_zip': self.input_reemplazo_zip.text().strip(),
            'sufijo_zip': self.input_sufijo_zip.text().strip(),
            'renombrar_directorios': self.check_activar_renombrado_directorios.isChecked(),
            'reemplazo_directorios': self.input_reemplazo_directorios.text().strip(),
            'sufijo_directorios': self.input_sufijo_directorios.text().strip(),
            'autorizacion_compensar': self.check_aut_compensar.isChecked(),
            'formato_json': self.check_formato_json.isChecked(),
            'txt_rips': self.check_txt_rips.isChecked(),
            'txt_docker': self.check_txt_docker.isChecked()
        }

        copiar_todo = self.check_copiar_sin_validar.isChecked()
        print(f"Modo copiar_todo: {copiar_todo}")

        if copiar_todo:
            respuesta = QMessageBox.question(self, "Confirmación", "Está copiando archivos sin validaciones, ¿desea continuar?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if respuesta == QMessageBox.No:
                print("Copia cancelada por el usuario.")
                return
            directorios_a_copiar = directorios_seleccionados
            print(f"Directorios a copiar (sin validar): {directorios_a_copiar}")
        else:
            directorios_a_copiar = directorios_seleccionados
            print(f"Directorios a copiar después de selección: {directorios_a_copiar}")

            if not directorios_a_copiar:
                QMessageBox.information(self, "Copia Cancelada", "No hay directorios válidos para copiar.")
                print("No hay directorios válidos después de la selección.")
                return

        texto_reemplazo_rips = self.input_reemplazo_rips.text().strip() if self.check_activar_renombrado_rips.isChecked() and not copiar_todo else ""
        texto_sufijo_rips = self.input_sufijo_rips.text().strip() if self.check_activar_renombrado_rips.isChecked() and not copiar_todo else ""
        texto_reemplazo_docker = self.input_reemplazo_docker.text().strip() if self.check_activar_renombrado_docker.isChecked() and not copiar_todo else ""
        texto_sufijo_docker = self.input_sufijo_docker.text().strip() if self.check_activar_renombrado_docker.isChecked() and not copiar_todo else ""
        texto_reemplazo_zip = self.input_reemplazo_zip.text().strip() if self.check_activar_renombrado_zip.isChecked() and not copiar_todo else ""
        texto_sufijo_zip = self.input_sufijo_zip.text().strip() if self.check_activar_renombrado_zip.isChecked() and not copiar_todo else ""
        texto_reemplazo_directorios = self.input_reemplazo_directorios.text().strip() if self.check_activar_renombrado_directorios.isChecked() and not copiar_todo else ""
        texto_sufijo_directorios = self.input_sufijo_directorios.text().strip() if self.check_activar_renombrado_directorios.isChecked() and not copiar_todo else ""

        self.boton_copiar.setEnabled(False)
        self.boton_cancelar.setEnabled(True)
        self.progress_bar.setValue(0)
        self.label_archivo_actual.setText("Archivo actual: ")
        self.inicio_copia = time.time()

        if hasattr(self, 'worker_thread') and self.worker_thread and self.worker_thread.isRunning():
            print("Esperando a que el hilo anterior termine...")
            self.worker_thread.quit()
            self.worker_thread.wait()

        self.worker_thread = WorkerCopia(
            directorios_a_copiar,
            directorio_destino,
            copiar_todo,
            opciones_copia,
            texto_reemplazo_rips,
            texto_sufijo_rips,
            texto_reemplazo_docker,
            texto_sufijo_docker,
            self.check_activar_Copiarenraiz.isChecked() and not copiar_todo,
            self.check_comprimir_zip.isChecked(),
            self.check_aut_compensar.isChecked(),
            texto_reemplazo_zip,
            texto_sufijo_zip,
            texto_reemplazo_directorios,
            texto_sufijo_directorios
        )

        self.worker_thread.progreso_signal.connect(self.actualizar_progreso_barra)
        self.worker_thread.log_signal.connect(self.log_copia)
        self.worker_thread.finalizado_signal.connect(self.copia_finalizada)
        self.worker_thread.error_signal.connect(self.mostrar_error)
        self.worker_thread.archivo_actual_signal.connect(self.actualizar_archivo_actual_label)
        self.worker_thread.notificacion_signal.connect(self.mostrar_notificacion)
        self.worker_thread.pregunta_signal.connect(self.manejar_pregunta)
        self.worker_thread.correccion_signal.connect(self.manejar_correccion)

        print("Hilo creado y señales conectadas, iniciando ejecución...")
        self.worker_thread.start()

    def copia_finalizada(self):
        tiempo_transcurrido = time.time() - self.inicio_copia if self.inicio_copia else 0
        self.boton_copiar.setEnabled(True)
        self.boton_cancelar.setEnabled(False)
        self.worker_thread.wait()

        if self.check_comprimir_zip.isChecked() and self.worker_thread.directorios_copiados:
            for nombre_directorio, archivos_copiados in self.worker_thread.directorios_copiados.items():
                if archivos_copiados:
                    temp_dir = os.path.dirname(archivos_copiados[0]) if not self.worker_thread.copiar_normales else None
                    if temp_dir and os.path.exists(temp_dir) and temp_dir.startswith(tempfile.gettempdir()):
                        shutil.rmtree(temp_dir)
                        self.log_text.append(f"Directorio temporal eliminado: {temp_dir}")

        self.worker_thread = None
        self.inicio_copia = None
        self.label_archivo_actual.setText("Copia finalizada")
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("Progreso: 100%")
        
        mensaje_final = f"Copia finalizada en {tiempo_transcurrido:.2f} segundos."
        if not hasattr(self.worker_thread, 'todos_los_errores') or not self.worker_thread.todos_los_errores:
            QMessageBox.information(self, "Copia Finalizada", mensaje_final)
        else:
            self.log_copia("Copia finalizada con problemas detectados, revisa el archivo de validación si se descargó.")
        
        self.log_copia(mensaje_final)

    def manejar_correccion(self, nombre_archivo, ruta_archivo, correcciones):
        mensaje = f"Se encontraron problemas en {nombre_archivo} con 'numAutorizacion'. ¿Desea corregir los siguientes valores?\n\n"
        with open(ruta_archivo, 'r', encoding='utf-8') as f:
            data = json.load(f)
        for path, corrected_value in correcciones.items():
            original_value = self.get_value_from_path(data, list(path))
            mensaje += f"- {'.'.join(map(str, path))}: \"{original_value}\" → \"{corrected_value}\"\n"
        mensaje += "\nSi elige 'Sí', el archivo será copiado y corregido en el destino."

        dialog = ResizableMessageDialog("Corrección de RIPS", mensaje, self)
        respuesta = dialog.exec_() == QDialog.Accepted
        self.worker_thread.respuesta_correccion = QMessageBox.Yes if respuesta else QMessageBox.No

    def get_value_from_path(self, data, path):
        current = data.get('usuarios', [{}])[0].get('servicios', {})
        for key in path:
            try:
                current = current[key]
            except (KeyError, IndexError, TypeError):
                return None
        return current

    def actualizar_contador_manual(self):
        texto_manual = self.manual_input_text.toPlainText()
        directorios = [dir.strip() for dir in texto_manual.strip().splitlines() if dir.strip()]
        contador = len(directorios)
        self.label_contador_manual.setText(f"Directorios: {contador}")

    def seleccionar_origen(self):
        directorio = QFileDialog.getExistingDirectory(self, "Seleccionar Directorio de Origen")
        if directorio:
            self.origen_input.setText(directorio)
            print(f"Directorio de origen seleccionado: {directorio}")

    def seleccionar_destino(self):
        directorio = QFileDialog.getExistingDirectory(self, "Seleccionar Directorio de Destino")
        if directorio:
            self.destino_input.setText(directorio)

    def actualizar_origen(self, texto):
        self.settings.setValue("ruta_origen", texto)
        self.boton_copiar.setEnabled(
            self.origen_input.text() != "Directorio de Origen:" and 
            self.destino_input.text() != "Directorio de Destino:" and 
            self.lista_directorios_copiar.count() > 0
        )
        print(f"Ruta de origen actualizada: {texto}")

    def actualizar_destino(self, texto):
        self.settings.setValue("ruta_destino", texto)
        self.boton_copiar.setEnabled(
            self.origen_input.text() != "Directorio de Origen:" and 
            self.destino_input.text() != "Directorio de Destino:" and 
            self.lista_directorios_copiar.count() > 0
        )
        print(f"Ruta de destino actualizada: {texto}")

    def actualizar_estado_renombrado(self):
        self.actualizar_estado_renombrado_rips()
        self.actualizar_estado_renombrado_docker()
        self.actualizar_estado_renombrado_zip()
        self.actualizar_estado_renombrado_directorios()

    def actualizar_estado_renombrado_rips(self):
        esta_activado = self.check_activar_renombrado_rips.isChecked()
        self.input_reemplazo_rips.setEnabled(esta_activado)
        self.input_sufijo_rips.setEnabled(esta_activado)

    def actualizar_estado_renombrado_docker(self):
        esta_activado = self.check_activar_renombrado_docker.isChecked()
        self.input_reemplazo_docker.setEnabled(esta_activado)
        self.input_sufijo_docker.setEnabled(esta_activado)

    def actualizar_estado_renombrado_zip(self):
        esta_activado = self.check_comprimir_zip.isChecked() and self.check_activar_renombrado_zip.isChecked()
        self.input_reemplazo_zip.setEnabled(esta_activado)
        self.input_sufijo_zip.setEnabled(esta_activado)

    def actualizar_estado_renombrado_directorios(self):
        enabled = self.check_activar_renombrado_directorios.isChecked()
        self.input_reemplazo_directorios.setEnabled(enabled)
        self.input_sufijo_directorios.setEnabled(enabled)
        self.settings.setValue("check_renombrado_directorios", self.check_activar_renombrado_directorios.isChecked())

    def agregar_listado(self):
        archivo, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo de Listado", "", "Archivos de Listado (*.txt *.xlsx *.xls);;Text Files (*.txt);;Excel Files (*.xlsx *.xls);;All Files (*)")
        if archivo:
            self.procesar_archivo_listado(archivo)
            self.boton_copiar.setEnabled(True)

    def agregar_listado_manual(self):
        directorio_origen = self.origen_input.text().strip()
        if not directorio_origen:
            QMessageBox.warning(self, "Advertencia", "Primero seleccione el directorio de origen.")
            print("Advertencia: No se ha seleccionado directorio de origen.")
            return

        print("Iniciando agregar_listado_manual...")
        directorios_no_validos = []
        directorios_manuales = self.manual_input_text.toPlainText().strip().splitlines()
        directorios_validos_contar = 0

        restringir_2025 = self.check_2025.isChecked()
        directorios_encontrados = {}

        if restringir_2025:
            print("Búsqueda restringida a carpetas '2025*/FACTURAS_SALUD'...")
            for carpeta in os.listdir(directorio_origen):
                if carpeta.startswith('2025'):
                    ruta_2025 = os.path.join(directorio_origen, carpeta)
                    if os.path.isdir(ruta_2025):
                        ruta_facturas_salud = os.path.join(ruta_2025, 'FACTURAS_SALUD')
                        if os.path.isdir(ruta_facturas_salud):
                            print(f"Explorando: {ruta_facturas_salud}")
                            for dir_name in os.listdir(ruta_facturas_salud):
                                ruta_completa = os.path.join(ruta_facturas_salud, dir_name)
                                if os.path.isdir(ruta_completa):
                                    directorios_encontrados[dir_name] = ruta_completa
                                    print(f"Encontrado: {ruta_completa}")
                        else:
                            print(f"No se encontró FACTURAS_SALUD en: {ruta_2025}")
                    else:
                        print(f"No es un directorio: {ruta_2025}")
        else:
            print("Buscando en todas las subcarpetas del directorio de origen...")
            for root, dirs, _ in os.walk(directorio_origen):
                for dir_name in dirs:
                    ruta_completa = os.path.join(root, dir_name)
                    directorios_encontrados[dir_name] = ruta_completa
                    print(f"Encontrado: {ruta_completa}")

        for directorio in directorios_manuales:
            directorio = directorio.strip()
            if not directorio:
                continue

            if directorio in directorios_encontrados:
                item = QListWidgetItem(directorio)
                item.setData(Qt.UserRole, directorios_encontrados[directorio])
                if item.text() not in [self.lista_directorios_copiar.item(i).text() for i in range(self.lista_directorios_copiar.count())]:
                    self.lista_directorios_copiar.addItem(item)
                    item.setSelected(True)  # Seleccionar automáticamente el ítem
                    directorios_validos_contar += 1
                    print(f"Agregado y seleccionado: {directorio} ({directorios_encontrados[directorio]})")
            else:
                directorios_no_validos.append(directorio)
                print(f"No encontrado: {directorio}")

        self.boton_copiar.setEnabled(self.destino_input.text().strip() != "" and self.lista_directorios_copiar.count() > 0)
        self.log_text.append(f"Directorios válidos agregados y seleccionados: {directorios_validos_contar}")

        if directorios_no_validos:
            mensaje = (f"Se agregaron {directorios_validos_contar} directorios válidos.\n\n"
                       "Los siguientes directorios no se encontraron en la ruta de origen:\n" +
                       "\n".join(f"- {d}" for d in directorios_no_validos) +
                       "\n\n¿Desea continuar con los directorios encontrados?")
            dialogo = ResizableMessageDialog("Directorios No Encontrados", mensaje, self)
            resultado = dialogo.exec_()
            if resultado == QDialog.Accepted:
                self.log_text.append("Usuario eligió continuar con los directorios encontrados.")
            else:
                self.log_text.append("Usuario canceló la operación debido a directorios no encontrados.")
                # Limpiar los directorios agregados si el usuario no quiere continuar
                for i in range(self.lista_directorios_copiar.count() - 1, -1, -1):
                    item = self.lista_directorios_copiar.item(i)
                    if item.text() in [d for d in directorios_manuales if d.strip()]:
                        self.lista_directorios_copiar.takeItem(i)
                self.boton_copiar.setEnabled(self.destino_input.text().strip() != "" and self.lista_directorios_copiar.count() > 0)
                return
        else:
            self.log_text.append("Todos los directorios manuales fueron encontrados y seleccionados.")

    def procesar_archivo_listado(self, archivo):
        directorio_origen = self.origen_input.text().strip()
        if not directorio_origen:
            QMessageBox.warning(self, "Advertencia", "Primero seleccione el directorio de origen.")
            return

        extension = os.path.splitext(archivo)[1].lower()
        directorios = []
        directorios_no_validos = []

        if extension == '.txt':
            with open(archivo, 'r', encoding='utf-8') as f:
                directorios = [linea.strip() for linea in f.readlines() if linea.strip()]
        elif extension in ['.xlsx', '.xls']:
            wb = openpyxl.load_workbook(archivo)
            ws = wb.active
            directorios = [str(cell.value).strip() for cell in ws['A'] if cell.value and str(cell.value).strip()]
            wb.close()

        if not directorios:
            QMessageBox.warning(self, "Advertencia", "El archivo seleccionado no contiene directorios válidos.")
            return

        restringir_2025 = self.check_2025.isChecked()
        directorios_encontrados = {}

        if restringir_2025:
            for carpeta in os.listdir(directorio_origen):
                if carpeta.startswith('2025'):
                    ruta_2025 = os.path.join(directorio_origen, carpeta)
                    if os.path.isdir(ruta_2025):
                        ruta_facturas_salud = os.path.join(ruta_2025, 'FACTURAS_SALUD')
                        if os.path.isdir(ruta_facturas_salud):
                            for dir_name in os.listdir(ruta_facturas_salud):
                                ruta_completa = os.path.join(ruta_facturas_salud, dir_name)
                                if os.path.isdir(ruta_completa):
                                    directorios_encontrados[dir_name] = ruta_completa
        else:
            for root, dirs, _ in os.walk(directorio_origen):
                for dir_name in dirs:
                    ruta_completa = os.path.join(root, dir_name)
                    directorios_encontrados[dir_name] = ruta_completa

        directorios_validos_contar = 0
        for directorio in directorios:
            if directorio in directorios_encontrados:
                item = QListWidgetItem(directorio)
                item.setData(Qt.UserRole, directorios_encontrados[directorio])
                if item.text() not in [self.lista_directorios_copiar.item(i).text() for i in range(self.lista_directorios_copiar.count())]:
                    self.lista_directorios_copiar.addItem(item)
                    item.setSelected(True)  # Seleccionar automáticamente el ítem
                    directorios_validos_contar += 1
            else:
                directorios_no_validos.append(directorio)

        self.log_text.append(f"Directorios válidos agregados y seleccionados desde archivo: {directorios_validos_contar}")
        if directorios_no_validos:
            mensaje = f"Se agregaron {directorios_validos_contar} directorios válidos.\n\nLos siguientes directorios no se encontraron:\n" + "\n".join(f"- {d}" for d in directorios_no_validos)
            QMessageBox.information(self, "Resultado de Carga de Archivo", mensaje)

        self.boton_copiar.setEnabled(self.destino_input.text().strip() != "" and self.lista_directorios_copiar.count() > 0)

    def eliminar_listado(self):
        items_seleccionados = self.lista_directorios_copiar.selectedItems()
        if not items_seleccionados:
            QMessageBox.warning(self, "Advertencia", "Seleccione al menos un directorio para eliminar.")
            return
        for item in items_seleccionados:
            self.lista_directorios_copiar.takeItem(self.lista_directorios_copiar.row(item))
        self.boton_copiar.setEnabled(self.destino_input.text().strip() != "" and self.lista_directorios_copiar.count() > 0)
        self.log_text.append(f"Eliminados {len(items_seleccionados)} directorios de la lista.")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        directorios = [url.toLocalFile() for url in urls if os.path.isdir(url.toLocalFile())]
        directorios_existentes = {self.lista_directorios_copiar.item(i).text() for i in range(self.lista_directorios_copiar.count())}
        
        for directorio in directorios:
            nombre_directorio = os.path.basename(directorio)
            if nombre_directorio not in directorios_existentes:
                item = QListWidgetItem(nombre_directorio)
                item.setData(Qt.UserRole, directorio)
                self.lista_directorios_copiar.addItem(item)
        
        self.boton_copiar.setEnabled(self.destino_input.text().strip() != "" and self.lista_directorios_copiar.count() > 0)
        self.log_text.append(f"Agregados {len(directorios)} directorios mediante arrastrar y soltar.")

    def actualizar_progreso_barra(self, valor):
        self.progress_bar.setValue(valor)
        self.progress_bar.setFormat(f"Progreso: {valor}%")

    def log_copia(self, mensaje):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"[{timestamp}] {mensaje}")

    def mostrar_error(self, mensaje):
        QMessageBox.critical(self, "Error", mensaje)
        self.boton_copiar.setEnabled(True)
        self.boton_cancelar.setEnabled(False)

    def actualizar_archivo_actual_label(self, archivo):
        self.label_archivo_actual.setText(f"Archivo actual: {archivo}")

    def mostrar_notificacion(self, mensaje):
        QMessageBox.information(self, "Notificación", mensaje)

    def manejar_pregunta(self, mensaje, botones):
        dialogo = ResizableMessageDialog("Confirmación", mensaje, self)
        resultado = dialogo.exec_()
        self.worker_thread.respuesta_usuario = QMessageBox.Yes if resultado == QDialog.Accepted else QMessageBox.No

    def cancelar_copia(self):
        if self.worker_thread and self.worker_thread.isRunning():
            respuesta = QMessageBox.question(self, "Cancelar Copia", "¿Está seguro de que desea cancelar la copia en curso?", QMessageBox.Yes | QMessageBox.No)
            if respuesta == QMessageBox.Yes:
                self.worker_thread.cancelar_flag = True
                self.worker_thread.quit()
                self.worker_thread.wait()
                self.boton_copiar.setEnabled(True)
                self.boton_cancelar.setEnabled(False)
                self.progress_bar.setValue(0)
                self.label_archivo_actual.setText("Copia cancelada")
                self.log_text.append("Copia cancelada por el usuario.")
                self.worker_thread = None

    def cambiar_tema(self):
        if self.check_tema_oscuro.isChecked():
            QApplication.instance().setStyleSheet("")
            self.check_tema_oscuro.setText("Tema Claro")
        else:
            QApplication.instance().setStyleSheet("""
                QWidget {
                    background-color: #2b2b2b;
                    color: #ffffff;
                }
                QPushButton {
                    background-color: #3c3f41;
                    border: 1px solid #555555;
                    padding: 5px;
                }
                QPushButton:hover {
                    background-color: #4a4e50;
                }
                QLineEdit, QTextEdit, QListWidget {
                    background-color: #353535;
                    border: 1px solid #555555;
                    color: #ffffff;
                }
                QCheckBox::indicator {
                    border: 1px solid #555555;
                    background-color: #353535;
                }
                QCheckBox::indicator:checked {
                    background-color: #4a90e2;
                }
                QProgressBar {
                    border: 1px solid #555555;
                    background-color: #353535;
                    text-align: center;
                }
                QProgressBar::chunk {
                    background-color: #4a90e2;
                }
                QGroupBox {
                    border: 1px solid #555555;
                    margin-top: 1em;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    padding: 0 3px;
                }
            """)
            self.check_tema_oscuro.setText("Tema Oscuro")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CopiadorDirectorios()
    window.show()
    sys.exit(app.exec_())