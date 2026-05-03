from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton,
    QLabel, QFileDialog, QCheckBox,
    QProgressBar, QTextEdit, QListWidget,
    QListWidgetItem, QMessageBox, QLineEdit,
    QGroupBox, QScrollArea, QWidget, QVBoxLayout,
    QDialog, QHBoxLayout, QFormLayout, QGridLayout,QProgressDialog,
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
import csv

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

    def __init__(self, directorios_origen, directorio_destino, copiar_todo, opciones_copia,
                  texto_reemplazo_rips, texto_sufijo_rips, texto_reemplazo_docker, texto_sufijo_docker,
                    copiar_en_raiz, comprimir_zip, aut_compensar, 
                    texto_reemplazo_zip, texto_sufijo_zip, texto_reemplazo_directorios, texto_sufijo_directorios, texto_reemplazo_ad, texto_sufijo_ad, descargar_cuv):
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
        self.texto_reemplazo_ad = texto_reemplazo_ad
        self.texto_sufijo_ad = texto_sufijo_ad
        self.descargar_cuv = descargar_cuv
        self.formato_json = opciones_copia.get('formato_json', False)
        self.cancelar_flag = False
        self.archivos_copiados = 0
        self.total_archivos = 0
        self.directorios_copiados = {}
        self.archivos_corregidos_copiados = set()
        self.todos_los_errores = []
        self.validacion_autorizaciones = {}
        self.respuesta_pregunta = {}
        self.cuv_data = []
        

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
            self.log_signal.emit(f"DEBUG CUV: descargar_cuv = {self.descargar_cuv}")
            self.log_signal.emit(f"DEBUG CUV: Cantidad de directorios a procesar = {len(self.directorios_origen)}")
            self.log_signal.emit("Iniciando copia de directorios...")
            self.log_signal.emit(f"Directorios origen iniciales: {self.directorios_origen}")
            self.preparar_conteo_archivos()

            directorios_a_copiar = self.directorios_origen
            self.todos_los_errores = []

            if self.copiar_todo:
                self.log_signal.emit("Copiando todos los directorios sin validaciones...")
                directorios_a_copiar_final = directorios_a_copiar
            else:
                problemas_rips = []
                directorios_problematicos_rips = set()
                directorios_problematicos_docker = set()
                directorios_problematicos_archivos = set()
                problemas_docker = []  # Acumula problemas de ResultState y existencia de archivos Docker

                # Validación de RIPS
                for dir_origen in self.directorios_origen:
                    rips_path = os.path.join(dir_origen, "RIPS")
                    nombre_directorio = os.path.basename(dir_origen)
                    rips_existe = os.path.isdir(rips_path)
                    self.log_signal.emit(f"Validando RIPS en {dir_origen} - Existe: {rips_existe}")
                    if not rips_existe:
                        error_msg = f"- '{nombre_directorio}': No se encontró la carpeta 'RIPS'"
                        problemas_rips.append(error_msg)
                        directorios_problematicos_rips.add(dir_origen)

                if problemas_rips:
                    mensaje_rips = "Los siguientes directorios no cuentan con la carpeta 'RIPS':\n" + "\n".join(problemas_rips) + "\n\n¿Desea copiar estos directorios?"
                    self.respuesta_usuario = None
                    self.pregunta_signal.emit(mensaje_rips, QMessageBox.Yes | QMessageBox.No)
                    while self.respuesta_usuario is None and not self.cancelar_flag:
                        time.sleep(0.1)
                    if self.respuesta_usuario == QMessageBox.No:
                        directorios_a_copiar = [d for d in directorios_a_copiar if d not in directorios_problematicos_rips]
                        self.log_signal.emit(f"Directorios omitidos por RIPS: {directorios_problematicos_rips}")
                    self.todos_los_errores.extend(problemas_rips)

                # Validación de ResultState y existencia de archivos Docker
                directorios_a_validar = directorios_a_copiar[:]
                for dir_origen in directorios_a_validar:
                    nombre_directorio = os.path.basename(dir_origen)
                    docker_files = [
                        os.path.join(root, file)
                        for root, _, files in os.walk(dir_origen)
                        for file in files
                        if file.startswith("ResultadosDoker_") and file.endswith(".json")
                    ]
                    self.log_signal.emit(f"Archivos Docker en {dir_origen}: {docker_files}")

                    # Nueva validación: si 'rips_docker' está activado y no hay archivos Docker
                    if self.opciones_copia['rips_docker'] and not docker_files:
                        error_msg = f"- '{nombre_directorio}': No se encontró ningún archivo 'ResultadosDoker_*.json'"
                        problemas_docker.append(error_msg)
                        directorios_problematicos_docker.add(dir_origen)
                        self.todos_los_errores.append(error_msg)
                        self.log_signal.emit(f"Directorio {nombre_directorio} agregado a problemas de Docker por falta de archivo")
                        continue  # Saltar a la siguiente iteración si no hay archivos

                    # Validación de ResultState (solo si hay archivos Docker)
                    result_state_encontrado = False
                    for ruta_docker in docker_files:
                        data = None
                        for encoding in ['utf-8', 'latin-1', 'windows-1252']:
                            try:
                                with open(ruta_docker, 'r', encoding=encoding) as f:
                                    data = json.load(f)
                                self.log_signal.emit(f"{ruta_docker} leído con codificación {encoding}")
                                break
                            except UnicodeDecodeError:
                                continue
                            except json.JSONDecodeError:
                                self.log_signal.emit(f"{ruta_docker} no se pudo leer como JSON")
                                break
                        if data is None:
                            self.log_signal.emit(f"{ruta_docker}: No se pudo decodificar con ninguna codificación soportada")
                            continue

                        result_state_key = next((key for key in data.keys() if key.lower() == "resultstate"), None)
                        if result_state_key and data[result_state_key]:
                            result_state_encontrado = True
                            self.log_signal.emit(f"{ruta_docker} tiene ResultState: true")
                            break

                    if not result_state_encontrado and docker_files:  # Solo si hay archivos pero no ResultState válido
                        error_msg = f"- '{nombre_directorio}': Tiene ResultState false, no tiene ResultState o no se pudo leer ningún archivo Docker"
                        problemas_docker.append(error_msg)
                        directorios_problematicos_docker.add(dir_origen)
                        self.todos_los_errores.append(error_msg)
                        self.log_signal.emit(f"Directorio {nombre_directorio} agregado a problemas de Docker por ResultState")

                # Pregunta consolidada para problemas de Docker
                if problemas_docker:
                    mensaje_docker = "Los siguientes directorios tienen problemas con archivos Docker:\n" + "\n".join(problemas_docker) + "\n\n¿Desea copiar estos directorios?"
                    self.respuesta_usuario = None
                    self.pregunta_signal.emit(mensaje_docker, QMessageBox.Yes | QMessageBox.No)
                    while self.respuesta_usuario is None and not self.cancelar_flag:
                        time.sleep(0.1)
                    if self.respuesta_usuario == QMessageBox.No:
                        directorios_a_copiar = [d for d in directorios_a_copiar if d not in directorios_problematicos_docker]
                        self.log_signal.emit(f"Directorios omitidos por Docker: {directorios_problematicos_docker}")
                    else:
                        self.log_signal.emit("Usuario eligió copiar directorios con problemas de Docker")

                # Validación de archivos AD, AR, FV por directorio
                directorios_a_validar = directorios_a_copiar[:]
                for dir_origen in directorios_a_validar:
                    nombre_directorio = os.path.basename(dir_origen)
                    archivos_presentes = set(f.lower() for _, _, files in os.walk(dir_origen) for f in files)
                    self.log_signal.emit(f"Archivos en {dir_origen}: {archivos_presentes}")
                    archivos_faltantes = []

                    if self.opciones_copia['ad_xml'] and not any(f.startswith("ad") and f.endswith(".xml") for f in archivos_presentes):
                        archivos_faltantes.append("ad.xml")
                    if self.opciones_copia['ar_xml'] and not any(f.startswith("ar") and f.endswith(".xml") for f in archivos_presentes):
                        archivos_faltantes.append("ar.xml")
                    if self.opciones_copia['fv_pdf'] and not any(f.startswith("fv") and f.endswith(".pdf") for f in archivos_presentes):
                        archivos_faltantes.append("fv.pdf")
                    if self.opciones_copia['fv_xml'] and not any(f.startswith("fv") and f.endswith(".xml") for f in archivos_presentes):
                        archivos_faltantes.append("fv.xml")

                    if archivos_faltantes:
                        mensaje_archivos = f"No se encontraron los siguientes archivos en el directorio '{nombre_directorio}': {', '.join(archivos_faltantes)}. ¿Desea copiar los demás archivos?"
                        self.respuesta_usuario = None
                        self.pregunta_signal.emit(mensaje_archivos, QMessageBox.Yes | QMessageBox.No)
                        while self.respuesta_usuario is None and not self.cancelar_flag:
                            time.sleep(0.1)
                        if self.respuesta_usuario == QMessageBox.No:
                            directorios_a_copiar.remove(dir_origen)
                            directorios_problematicos_archivos.add(dir_origen)
                            error_msg = f"- '{nombre_directorio}': Faltan los archivos {', '.join(archivos_faltantes)} (omitido por usuario)"
                            self.todos_los_errores.append(error_msg)
                            self.log_signal.emit(f"Directorio {nombre_directorio} omitido por falta de archivos")
                        else:
                            error_msg = f"- '{nombre_directorio}': Faltan los archivos {', '.join(archivos_faltantes)} (copiado resto por usuario)"
                            self.todos_los_errores.append(error_msg)
                            self.log_signal.emit(f"Directorio {nombre_directorio} copiado a pesar de falta de archivos")

                directorios_a_copiar_final = set(directorios_a_copiar)
                self.log_signal.emit(f"Directorios a copiar tras validaciones: {directorios_a_copiar_final}")

            # Renombrado y copia de directorios
            directorios_con_renombrado = {}
            for dir_origen in directorios_a_copiar_final:
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

            self.validacion_autorizaciones = {}
            for dir_origen, dir_destino in directorios_con_renombrado.items():
                if self.comprimir_zip:
                    temp_dir = tempfile.mkdtemp()
                    archivos_copiados = self.copiar_archivos(dir_origen, temp_dir, self.copiar_todo, self.opciones_copia, self.copiar_en_raiz)
                    if archivos_copiados:
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
                    archivos_copiados = self.copiar_archivos(dir_origen, dir_destino, self.copiar_todo, self.opciones_copia, self.copiar_en_raiz)
                    if not hasattr(self, 'directorios_copiados'):
                        self.directorios_copiados = {}
                    self.directorios_copiados[dir_origen] = archivos_copiados

                if self.aut_compensar and os.path.basename(dir_origen) in self.validacion_autorizaciones:
                    self.todos_los_errores.extend(self.validacion_autorizaciones[os.path.basename(dir_origen)])

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
            self.error_signal.emit(f"Error durante la copia: {str(e)}")
            if self.descargar_cuv and self.cuv_data:
                try:
                    archivo_cuv = os.path.join(self.directorio_destino, "CUVs_Exportados.txt")
                    
                    with open(archivo_cuv, 'w', encoding='utf-8', newline='') as f:
                        writer = csv.writer(f, delimiter=',', quoting=csv.QUOTE_MINIMAL)
                        # Encabezados
                        writer.writerow(["NumFactura", "ProcesoId", "CodigoUnicoValidacion", 
                                       "ResultState", "CantidadValidaciones", "Validaciones"])
                        
                        # Datos
                        for fila in self.cuv_data:
                            writer.writerow(fila)
                    
                    self.log_signal.emit(f"✅ Archivo CUVs_Exportados.txt generado correctamente en el destino con {len(self.cuv_data)} facturas.")
                except Exception as e:
                    self.log_signal.emit(f"❌ Error al generar archivo CUV: {str(e)}")
            elif self.descargar_cuv:
                self.log_signal.emit("⚠️ No se encontraron archivos Docker válidos para generar CUV.")

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
                    
                    if file.startswith("ad") and file.endswith(".xml") and opciones_copia.get('renombrar_ad', False):
                        prefijo = self.texto_reemplazo_ad if self.texto_reemplazo_ad else ""
                        sufijo = self.texto_sufijo_ad if self.texto_sufijo_ad else ""
                        ext = os.path.splitext(file)[1]  # mantiene .xml
                        nombre_archivo_final = f"{prefijo}{nombre_directorio}{sufijo}{ext}"
                        self.log_signal.emit(f"DEBUG RENOMBRADO AD: prefijo='{prefijo}', sufijo='{sufijo}', base='{nombre_directorio}', resultado='{nombre_archivo_final}'")
                    
                    # Renombrado RIPS
                    elif es_rips and opciones_copia.get('renombrar_rips', False):
                        prefijo = self.texto_reemplazo_rips if self.texto_reemplazo_rips else ""
                        sufijo = self.texto_sufijo_rips if self.texto_sufijo_rips else ""
                        ext = '.txt' if txt_rips else os.path.splitext(file)[1]
                        nombre_archivo_final = f"{prefijo}{nombre_directorio}{sufijo}{ext}"
                        self.log_signal.emit(f"DEBUG RENOMBRADO RIPS: prefijo='{prefijo}', sufijo='{sufijo}', base='{nombre_directorio}', ext='{ext}'")
                    
                    # Renombrado Docker
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
                            if self.descargar_cuv and es_docker:
                                try:
                                    docker_data = None
                                    # Forzar recarga del JSON del Docker para mayor seguridad
                                    for encoding in ['utf-8', 'latin-1', 'windows-1252']:
                                        try:
                                            with open(ruta_origen, 'r', encoding=encoding) as f:
                                                docker_data = json.load(f)
                                            break
                                        except:
                                            continue
                                    
                                    if docker_data and isinstance(docker_data, dict):
                                        self.extraer_datos_cuv(docker_data, nombre_directorio)
                                    else:
                                        self.log_signal.emit(f"⚠️ No se pudo leer JSON de Docker: {file}")
                                except Exception as e:
                                    self.log_signal.emit(f"Error al extraer CUV de {file}: {str(e)}")

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
    
    def extraer_datos_cuv(self, data, nombre_directorio):
        """Extrae información del ResultadosDoker para el archivo CUV"""
        try:
            num_factura = str(data.get("NumFactura", ""))
            proceso_id = str(data.get("ProcesoId", ""))
            cuv = str(data.get("CodigoUnicoValidacion", ""))
            result_state = str(data.get("ResultState", False)).lower()

            # Procesar lista de validaciones
            resultados = data.get("ResultadosValidacion", [])
            cantidad = len(resultados)

            validaciones_list = []
            for res in resultados:
                clase = res.get("Clase", "")
                codigo = res.get("Codigo", "")
                desc = res.get("Descripcion", "")
                obs = res.get("Observaciones", "")
                
                texto = f"{codigo} - {clase} - {desc}"
                if obs:
                    texto += f" - {obs}"
                validaciones_list.append(texto)

            validaciones_concat = " | ".join(validaciones_list)

            self.cuv_data.append([
                num_factura,
                proceso_id,
                cuv,
                result_state,
                cantidad,
                validaciones_concat
            ])

            self.log_signal.emit(f"✓ CUV extraído: {num_factura} ({nombre_directorio}) - {cantidad} validaciones")

        except Exception as e:
            self.log_signal.emit(f"Error extrayendo CUV de {nombre_directorio}: {str(e)}")

class CacheWorker(QThread):
    progreso_signal = pyqtSignal(int)  # Para actualizar la barra de progreso
    finalizado_signal = pyqtSignal(dict)  # Para enviar el caché completo
    error_signal = pyqtSignal(str)  # Para manejar errores

    def __init__(self, directorio_origen, restringir_busqueda, anio, meses):
        super().__init__()
        self.directorio_origen = directorio_origen
        self.restringir_busqueda = restringir_busqueda
        self.anio = anio
        self.meses = meses

    def run(self):
        try:
            cache = {}
            if self.restringir_busqueda:
                print("Cacheando directorios restringidos por año y meses...")
                meses_lista = [m.strip() for m in self.meses.split(',') if m.strip()]
                total_carpetas = 0
                subdirectorios_totales = []

                if not self.anio.isdigit() or len(self.anio) != 4:
                    self.error_signal.emit(f"Error: El año '{self.anio}' no es válido. Usando '2025' por defecto.")
                    self.anio = "2025"

                # Primero contamos el total de subdirectorios para un progreso preciso
                for mes in meses_lista:
                    if not (len(mes) == 2 and mes.isdigit() and 1 <= int(mes) <= 12):
                        self.error_signal.emit(f"Error: El mes '{mes}' no es válido. Ignorando...")
                        continue
                    carpeta_mes = f"{self.anio}{mes}"
                    ruta_carpeta = os.path.join(self.directorio_origen, carpeta_mes)
                    if os.path.isdir(ruta_carpeta):
                        ruta_facturas_salud = os.path.join(ruta_carpeta, 'FACTURAS_SALUD')
                        if os.path.isdir(ruta_facturas_salud):
                            subdirectorios = [d for d in os.listdir(ruta_facturas_salud) if os.path.isdir(os.path.join(ruta_facturas_salud, d))]
                            subdirectorios_totales.extend([(mes, d) for d in subdirectorios])
                            total_carpetas += len(subdirectorios)

                progreso = 0
                for mes, dir_name in subdirectorios_totales:
                    carpeta_mes = f"{self.anio}{mes}"
                    ruta_facturas_salud = os.path.join(self.directorio_origen, carpeta_mes, 'FACTURAS_SALUD')
                    ruta_completa = os.path.join(ruta_facturas_salud, dir_name)
                    clave_cache = f"{carpeta_mes}_{dir_name}"
                    cache[clave_cache] = ruta_completa
                    progreso += 1
                    porcentaje = min(int((progreso / total_carpetas) * 100), 100) if total_carpetas > 0 else 100
                    self.progreso_signal.emit(porcentaje)
                    print(f"Cacheado: {ruta_completa} (Año: {self.anio}, Mes: {mes})")

            else:
                print("Cacheando solo los directorios en la raíz del directorio de origen...")
                items = os.listdir(self.directorio_origen)
                total_items = len([i for i in items if os.path.isdir(os.path.join(self.directorio_origen, i))])
                progreso = 0

                for item in items:
                    ruta_completa = os.path.join(self.directorio_origen, item)
                    if os.path.isdir(ruta_completa):
                        cache[item] = ruta_completa
                        progreso += 1
                        porcentaje = min(int((progreso / total_items) * 100), 100) if total_items > 0 else 100
                        self.progreso_signal.emit(porcentaje)
                        print(f"Cacheado (raíz): {ruta_completa}")

            # Asegurar que el progreso llegue al 100% antes de finalizar
            self.progreso_signal.emit(100)
            self.finalizado_signal.emit(cache)
        except Exception as e:
            self.error_signal.emit(f"Error al cargar el caché: {str(e)}")
            
class CopiadorDirectorios(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Copiador de Directorios")
        self.setMinimumSize(950, 950)  # Tamaño mínimo para la ventana
        self.setWindowIcon(QIcon(resource_path("icons/kurama.png")))
        self.settings = QSettings("MiEmpresa", "CopiadorDirectorios")

        self.directorios_cache = {}
        self.directorio_origen_cacheado = None
        self.cache_worker = None

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
        self.boton_origen.setIcon(QIcon(resource_path("icons/carpeta.png")))
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
        self.boton_destino.setIcon(QIcon(resource_path("icons/carpeta.png")))
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
        
        # Nueva sección para configurar año y meses
        self.check_restringir_busqueda = QCheckBox("Restringir Búsqueda")
        self.anio_input = QLineEdit("2025")
        self.anio_input.setPlaceholderText("Año (ej. 2025)")
        self.anio_input.setEnabled(False)
        self.meses_input = QLineEdit("02,03")
        self.meses_input.setPlaceholderText("Meses (ej. 02,03)")
        self.meses_input.setEnabled(False)
        self.boton_cargar_cache = QPushButton("Cargar Caché")
        self.boton_cargar_cache.clicked.connect(self.cargar_cache_manual)

        self.check_txt_rips = QCheckBox("txt Rips")
        self.check_txt_docker = QCheckBox("txt Docker")
        self.check_tema_oscuro = QCheckBox("Tema Claro")
        self.check_activar_Copiarenraiz = QCheckBox("Copiar en raíz")
        self.check_formato_json = QCheckBox("Formato JSON")
        self.check_comprimir_zip = QCheckBox("Comprimir ZIP")
        self.check_comprimir_zip.stateChanged.connect(self.actualizar_estado_comprimir)
        self.check_aut_compensar = QCheckBox("Aut Compensar")
        self.check_copiar_sin_validar = QCheckBox("Sin validar")
        self.check_descargar_cuv = QCheckBox("Descargar CUV")

        # Reorganizar checkboxes en una cuadrícula 3x3
        extra_options_layout.addWidget(self.check_restringir_busqueda, 0, 0)
        extra_options_layout.addWidget(QLabel("Año:"), 0, 1)
        extra_options_layout.addWidget(self.anio_input, 0, 2)
        extra_options_layout.addWidget(QLabel("Meses:"), 0, 3)
        extra_options_layout.addWidget(self.meses_input, 0, 4)
        extra_options_layout.addWidget(self.boton_cargar_cache, 0, 5)  # Botón en la misma fila
        extra_options_layout.addWidget(self.check_txt_rips, 1, 0)
        extra_options_layout.addWidget(self.check_txt_docker, 1, 1)
        extra_options_layout.addWidget(self.check_tema_oscuro, 1, 2)
        extra_options_layout.addWidget(self.check_activar_Copiarenraiz, 1, 3)
        extra_options_layout.addWidget(self.check_formato_json, 1, 4)
        extra_options_layout.addWidget(self.check_comprimir_zip, 2, 0)
        extra_options_layout.addWidget(self.check_aut_compensar, 2, 1)
        extra_options_layout.addWidget(self.check_copiar_sin_validar, 2, 2)
        extra_options_layout.addWidget(self.check_descargar_cuv, 2, 3)

        top_section_layout.addWidget(extra_options_group)
        top_section_layout.setStretch(1, 2) # Columna derecha ocupa 2/3 del espacio

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

        # renombrar ad 

        self.grupo_renombrado_ad = QGroupBox("Renombrado AD.XML")
        ad_layout = QFormLayout()
        self.grupo_renombrado_ad.setLayout(ad_layout)
        
        self.check_activar_renombrado_ad = QCheckBox("Renombrar AD.XML")
        self.check_activar_renombrado_ad.stateChanged.connect(self.actualizar_estado_renombrado_ad)
        
        self.input_reemplazo_ad = QLineEdit()
        self.input_reemplazo_ad.setPlaceholderText("Opcional: Prefijo para AD")
        self.input_sufijo_ad = QLineEdit()
        self.input_sufijo_ad.setPlaceholderText("Opcional: Sufijo para AD")
        
        ad_layout.addRow(self.check_activar_renombrado_ad)
        ad_layout.addRow(QLabel("Prefijo:"), self.input_reemplazo_ad)
        ad_layout.addRow(QLabel("Sufijo:"), self.input_sufijo_ad)
        
        renaming_section_layout.addWidget(self.grupo_renombrado_ad)

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
        self.boton_agregar_listado.setIcon(QIcon(resource_path("icons/nota.png")))
        self.boton_agregar_listado.clicked.connect(self.agregar_listado)
        self.boton_eliminar_listado = QPushButton("Eliminar Listado")
        self.boton_eliminar_listado.setIcon(QIcon(resource_path("icons/eliminar.png")))
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
        self.boton_agregar_manual.setIcon(QIcon(resource_path("icons/editar.png")))
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
        self.boton_copiar.setIcon(QIcon(resource_path("icons/copiar.png")))
        self.boton_copiar.clicked.connect(self.confirmar_copia)
        self.boton_copiar.setEnabled(False)
        self.boton_cancelar = QPushButton("Cancelar")
        self.boton_cancelar.setIcon(QIcon(resource_path("icons/cancelar.png")))
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
        self.check_restringir_busqueda.stateChanged.connect(self.toggle_restringir_busqueda)
        self.setAcceptDrops(True)
        self.cargar_configuraciones() # Asegurarnos de cargar configuraciones iniciales
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
        self.toggle_restringir_busqueda() 

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
        self.check_restringir_busqueda.setChecked(self.settings.value("restringir_busqueda", False, type=bool))
        self.anio_input.setText(self.settings.value("anio_restriccion", "2025", type=str))
        self.meses_input.setText(self.settings.value("meses_restriccion", "02,03", type=str))
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
        self.check_activar_renombrado_ad.setChecked(self.settings.value("check_renombrado_ad", False, type=bool))
        self.input_reemplazo_ad.setText(self.settings.value("reemplazo_ad", "", type=str))
        self.input_sufijo_ad.setText(self.settings.value("sufijo_ad", "", type=str))
        self.check_descargar_cuv.setChecked(self.settings.value("check_descargar_cuv", False, type=bool))

        self.boton_copiar.setEnabled(
            self.origen_input.text().strip() != "" and 
            self.destino_input.text().strip() != "" and 
            self.lista_directorios_copiar.count() > 0
        )
        self.toggle_restringir_busqueda()

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

    def toggle_restringir_busqueda(self):
        """Activa o desactiva los campos Año y Meses, y habilita/deshabilita el botón Cargar Caché."""
        estado = self.check_restringir_busqueda.isChecked()
        print(f"Toggle restringir búsqueda: estado={estado}")
        self.anio_input.setEnabled(estado)
        self.meses_input.setEnabled(estado)
        self.boton_cargar_cache.setEnabled(estado)

        if not estado:
            self.log_text.append("Opción 'Restringir Búsqueda' desactivada. Los campos Año y Meses han sido deshabilitados.")
        else:
            self.log_text.append("Opción 'Restringir Búsqueda' activada. Ahora puedes especificar Año y Meses y cargar la caché manualmente.")

        print(f"Año habilitado: {self.anio_input.isEnabled()}, Meses habilitado: {self.meses_input.isEnabled()}, Botón Cargar Caché habilitado: {self.boton_cargar_cache.isEnabled()}")
        self.update()
        
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
        self.settings.setValue("anio_restriccion", self.anio_input.text())
        self.settings.setValue("meses_restriccion", self.meses_input.text())
        self.settings.setValue("check_renombrado_ad", self.check_activar_renombrado_ad.isChecked())
        self.settings.setValue("reemplazo_ad", self.input_reemplazo_ad.text())
        self.settings.setValue("sufijo_ad", self.input_sufijo_ad.text())
        self.settings.setValue("check_descargar_cuv", self.check_descargar_cuv.isChecked())

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
            'txt_docker': self.check_txt_docker.isChecked(),
            'renombrar_ad': self.check_activar_renombrado_ad.isChecked(),
            'reemplazo_ad': self.input_reemplazo_ad.text().strip(),
            'sufijo_ad': self.input_sufijo_ad.text().strip(),
            'descargar_cuv': self.check_descargar_cuv.isChecked(),
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
        texto_reemplazo_ad = self.input_reemplazo_ad.text().strip() if self.check_activar_renombrado_ad.isChecked() and not copiar_todo else ""
        texto_sufijo_ad = self.input_sufijo_ad.text().strip() if self.check_activar_renombrado_ad.isChecked() and not copiar_todo else ""

        self.boton_copiar.setEnabled(False)
        self.boton_cancelar.setEnabled(True)
        self.progress_bar.setValue(0)
        self.label_archivo_actual.setText("Archivo actual: ")
        self.inicio_copia = time.time()

        if hasattr(self, 'worker_thread') and self.worker_thread and self.worker_thread.isRunning():
            print("Esperando a que el hilo anterior termine...")
            self.worker_thread.quit()
            self.worker_thread.wait()

        print("Preparando WorkerCopia con Descargar CUV:", self.check_descargar_cuv.isChecked())    

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
            texto_sufijo_directorios,
            texto_reemplazo_ad,
            texto_sufijo_ad,
            self.check_descargar_cuv.isChecked()
        )
        print("WorkerCopia creado correctamente con descargar_cuv =", self.check_descargar_cuv.isChecked())

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
            # No cargar el caché aquí; el usuario debe usar "Cargar Caché"

    def seleccionar_destino(self):
        directorio = QFileDialog.getExistingDirectory(self, "Seleccionar Directorio de Destino")
        if directorio:
            self.destino_input.setText(directorio)

    def cargar_cache_manual(self):
        """Carga el caché manualmente según los parámetros actuales, solo si "Restringir Búsqueda" está activado."""
        if not self.check_restringir_busqueda.isChecked():
            QMessageBox.warning(self, "Advertencia", "La opción 'Restringir Búsqueda' debe estar activada para cargar el caché.")
            return

        directorio_origen = self.origen_input.text().strip()
        if not directorio_origen or not os.path.isdir(directorio_origen):
            QMessageBox.warning(self, "Advertencia", "Por favor, seleccione un directorio de origen válido antes de cargar el caché.")
            return
        print(f"Cargando caché manualmente para: {directorio_origen}")
        self.cargar_cache_con_progreso(directorio_origen)

    def actualizar_cache_directorios(self, directorio_origen):
        """Cachea los directorios en la ruta de origen basándose en año y meses configurables."""
        self.directorios_cache.clear()
        self.directorio_origen_cacheado = directorio_origen
        restringir_2025 = self.check_2025.isChecked()

        if restringir_2025:
            print("Cacheando directorios restringidos por año y meses...")
            anio = self.anio_input.text().strip() or "2025"  # Valor por defecto si está vacío
            meses_str = self.meses_input.text().strip() or "01,12"  # Valor por defecto si está vacío
            meses = [m.strip() for m in meses_str.split(',') if m.strip()]  # Convertir a lista de meses

            if not anio.isdigit() or len(anio) != 4:
                self.log_text.append(f"Error: El año '{anio}' no es válido. Usando '2025' por defecto.")
                anio = "2025"

            for mes in meses:
                if not (len(mes) == 2 and mes.isdigit() and 1 <= int(mes) <= 12):
                    self.log_text.append(f"Error: El mes '{mes}' no es válido. Ignorando...")
                    continue

                # Formar el nombre de la carpeta (ej. "202502", "202503")
                carpeta_mes = f"{anio}{mes}"
                ruta_carpeta = os.path.join(directorio_origen, carpeta_mes)

                if os.path.isdir(ruta_carpeta):
                    ruta_facturas_salud = os.path.join(ruta_carpeta, 'FACTURAS_SALUD')
                    if os.path.isdir(ruta_facturas_salud):
                        for dir_name in os.listdir(ruta_facturas_salud):
                            ruta_completa = os.path.join(ruta_facturas_salud, dir_name)
                            if os.path.isdir(ruta_completa):
                                clave_cache = f"{carpeta_mes}_{dir_name}"  # Clave única
                                self.directorios_cache[clave_cache] = ruta_completa
                                print(f"Cacheado: {ruta_completa} (Año: {anio}, Mes: {mes})")
                    else:
                        print(f"No se encontró FACTURAS_SALUD en: {ruta_carpeta}")
                else:
                    print(f"No se encontró la carpeta {carpeta_mes} en: {directorio_origen}")
        else:
            print("Cacheando solo los directorios en la raíz del directorio de origen...")
            # Solo buscar en la raíz del directorio de origen
            for item in os.listdir(directorio_origen):
                ruta_completa = os.path.join(directorio_origen, item)
                if os.path.isdir(ruta_completa):  # Solo agregar si es un directorio
                    self.directorios_cache[item] = ruta_completa
                    print(f"Cacheado (raíz): {ruta_completa}")

    def cargar_cache_con_progreso(self, directorio_origen):
        """Carga el caché con una barra de progreso, solo si "Restringir Búsqueda" está activado."""
        if not self.check_restringir_busqueda.isChecked():
            self.log_text.append("No se cargó el caché porque 'Restringir Búsqueda' está desactivado.")
            return

        if self.cache_worker and self.cache_worker.isRunning():
            self.cache_worker.quit()
            self.cache_worker.wait()

        progress_dialog = QProgressDialog("Cargando caché...", "Cancelar", 0, 100, self)
        progress_dialog.setWindowTitle("Progreso de Caché")
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.setAutoClose(False)
        progress_dialog.setValue(0)

        self.cache_worker = CacheWorker(
            directorio_origen,
            self.check_restringir_busqueda.isChecked(),
            self.anio_input.text().strip() or "2025",
            self.meses_input.text().strip() or "01,12"
        )

        self.cache_worker.progreso_signal.connect(progress_dialog.setValue)
        self.cache_worker.finalizado_signal.connect(lambda cache: self.on_cache_finalizado(cache, progress_dialog))
        self.cache_worker.error_signal.connect(lambda msg: self.on_cache_error(msg, progress_dialog))
        progress_dialog.canceled.connect(self.cancelar_cacheo)

        self.cache_worker.start()
        progress_dialog.exec_()

    def on_cache_finalizado(self, cache, progress_dialog):
        """Maneja la finalización del cacheo."""
        self.directorios_cache = cache
        self.directorio_origen_cacheado = self.origen_input.text().strip()
        print(f"Caché cargado con {len(self.directorios_cache)} directorios.")
        self.log_text.append(f"Caché cargado para {self.directorio_origen_cacheado} con {len(self.directorios_cache)} directorios.")
        progress_dialog.setValue(100)  # Asegurar que llegue al 100%
        progress_dialog.close()  # Cerrar manualmente

    def on_cache_error(self, mensaje, progress_dialog):
        """Maneja errores durante el cacheo."""
        QMessageBox.critical(self, "Error", mensaje)
        self.directorios_cache.clear()
        self.directorio_origen_cacheado = None
        self.log_text.append(f"Error al cargar caché: {mensaje}")
        progress_dialog.close()

    def cancelar_cacheo(self):
        """Cancela el proceso de cacheo si el usuario presiona Cancelar."""
        if self.cache_worker and self.cache_worker.isRunning():
            self.cache_worker.quit()
            self.cache_worker.wait()
            self.directorios_cache.clear()
            self.directorio_origen_cacheado = None
            self.log_text.append("Carga del caché cancelada por el usuario.")

    def buscar_directorio_en_cache(self, nombre_directorio):
        """Busca un directorio por su nombre, ya sea en la caché (si "Restringir Búsqueda" está activado) o directamente en la raíz del directorio de origen."""
        directorio_origen = self.origen_input.text().strip()

        if not directorio_origen or not os.path.isdir(directorio_origen):
            return None

        if self.check_restringir_busqueda.isChecked():
            # Buscar en la caché si "Restringir Búsqueda" está activado
            if not self.directorios_cache:  # Asegúrate de que la caché exista
                return None
            anio = self.anio_input.text().strip() or "2025"
            meses_str = self.meses_input.text().strip() or "01,12"
            meses = [m.strip() for m in meses_str.split(',') if m.strip()]
            for mes in meses:
                clave = f"{anio}{mes}_{nombre_directorio}"
                if clave in self.directorios_cache:
                    return self.directorios_cache[clave]
            return None
        else:
            # Si "Restringir Búsqueda" está desactivado, buscar directamente en la raíz del directorio de origen
            ruta_completa = os.path.join(directorio_origen, nombre_directorio)
            if os.path.isdir(ruta_completa):
                return ruta_completa
            return None

    def actualizar_origen(self, texto):
        self.settings.setValue("ruta_origen", texto)
        self.boton_copiar.setEnabled(
            self.origen_input.text().strip() != "" and 
            self.destino_input.text().strip() != "" and 
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
        self.actualizar_estado_renombrado_ad()

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

    def actualizar_estado_renombrado_ad(self):
        esta_activado = self.check_activar_renombrado_ad.isChecked()
        self.input_reemplazo_ad.setEnabled(esta_activado)
        self.input_sufijo_ad.setEnabled(esta_activado)

    def agregar_listado(self):
        archivo, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo de Listado", "", "Archivos de Listado (*.txt *.xlsx *.xls);;Text Files (*.txt);;Excel Files (*.xlsx *.xls);;All Files (*)")
        if archivo:
            directorio_origen = self.origen_input.text().strip()
            if not directorio_origen:
                QMessageBox.warning(self, "Advertencia", "Primero seleccione un directorio de origen.")
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

            directorios_validos_contar = 0
            for directorio in directorios:
                ruta_completa = self.buscar_directorio_en_cache(directorio)  # Usa el nuevo método
                if ruta_completa:
                    item = QListWidgetItem(directorio)
                    item.setData(Qt.UserRole, ruta_completa)
                    if item.text() not in [self.lista_directorios_copiar.item(i).text() for i in range(self.lista_directorios_copiar.count())]:
                        self.lista_directorios_copiar.addItem(item)
                        item.setSelected(True)
                        directorios_validos_contar += 1
                else:
                    directorios_no_validos.append(directorio)

            self.log_text.append(f"Directorios válidos agregados y seleccionados desde archivo: {directorios_validos_contar}")
            if directorios_no_validos:
                mensaje = f"Se agregaron {directorios_validos_contar} directorios válidos.\n\nLos siguientes directorios no se encontraron:\n" + "\n".join(f"- {d}" for d in directorios_no_validos)
                QMessageBox.information(self, "Resultado de Carga de Archivo", mensaje)

            self.boton_copiar.setEnabled(self.destino_input.text().strip() != "" and self.lista_directorios_copiar.count() > 0)

    def agregar_listado_manual(self):
        directorio_origen = self.origen_input.text().strip()
        if not directorio_origen:
            QMessageBox.warning(self, "Advertencia", "Primero seleccione un directorio de origen.")
            print("Advertencia: No se ha seleccionado un directorio de origen.")
            return

        print("Iniciando agregar_listado_manual...")
        directorios_no_validos = []
        directorios_manuales = self.manual_input_text.toPlainText().strip().splitlines()
        directorios_validos_contar = 0

        for directorio in directorios_manuales:
            directorio = directorio.strip()
            if not directorio:
                continue

            ruta_completa = self.buscar_directorio_en_cache(directorio)  # Usa el nuevo método
            if ruta_completa:
                item = QListWidgetItem(directorio)
                item.setData(Qt.UserRole, ruta_completa)
                if item.text() not in [self.lista_directorios_copiar.item(i).text() for i in range(self.lista_directorios_copiar.count())]:
                    self.lista_directorios_copiar.addItem(item)
                    item.setSelected(True)
                    directorios_validos_contar += 1
                    print(f"Agregado y seleccionado: {directorio} ({ruta_completa})")
            else:
                directorios_no_validos.append(directorio)
                print(f"No encontrado en caché o raíz: {directorio}")

        self.boton_copiar.setEnabled(self.destino_input.text().strip() != "" and self.lista_directorios_copiar.count() > 0)
        self.log_text.append(f"Directorios válidos agregados y seleccionados: {directorios_validos_contar}")

        if directorios_no_validos:
            mensaje = (f"Se agregaron {directorios_validos_contar} directorios válidos.\n\n"
                    "Los siguientes directorios no se encontraron en el caché o raíz:\n" +
                    "\n".join(f"- {d}" for d in directorios_no_validos) +
                    "\n\n¿Desea continuar con los directorios encontrados?")
            dialogo = ResizableMessageDialog("Directorios No Encontrados", mensaje, self)
            resultado = dialogo.exec_()
            if resultado == QDialog.Accepted:
                self.log_text.append("Usuario eligió continuar con los directorios encontrados.")
            else:
                self.log_text.append("Usuario canceló la operación debido a directorios no encontrados.")
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
            QMessageBox.warning(self, "Advertencia", "Primero seleccione un directorio de origen.")
            return
        if directorio_origen != self.directorio_origen_cacheado:
            QMessageBox.warning(self, "Advertencia", "El caché no está actualizado para este directorio. Haga clic en 'Cargar Caché' primero.")
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

        directorios_validos_contar = 0
        for directorio in directorios:
            ruta_completa = self.buscar_directorio_en_cache(directorio)
            if ruta_completa:
                item = QListWidgetItem(directorio)
                item.setData(Qt.UserRole, ruta_completa)
                if item.text() not in [self.lista_directorios_copiar.item(i).text() for i in range(self.lista_directorios_copiar.count())]:
                    self.lista_directorios_copiar.addItem(item)
                    item.setSelected(True)
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
                ruta_completa = self.buscar_directorio_en_cache(nombre_directorio)  # Usa el nuevo método
                if ruta_completa:
                    item = QListWidgetItem(nombre_directorio)
                    item.setData(Qt.UserRole, ruta_completa)
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