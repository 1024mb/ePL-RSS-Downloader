#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
import re
import shutil
import stat
import subprocess
import sys
import xml.dom.minidom
from urllib import request

from unicodedata import normalize

try:
    from PyQt5 import QtGui, QtWidgets, uic
    from PyQt5.QtCore import Qt, QTimer, pyqtSlot, QTranslator, QLibraryInfo, QLocale, QThread
except:
    print('Es necesario tener instaladas las librerías gráficas Qt5 en Python para que funcione')
    sys.exit(1)

# Declaración de constantes
NUM_ENLACES = 300
VELOCIDAD = 5
VERSION = "3.5"
TORRENT_CLIENTE = ""
RSS_links = {"RSS semanal": "https://rss.epublibre.org/rss/semanal",
             "RSS mensual": "https://rss.epublibre.org/rss/mensual",
             "RSS total": "https://rss.epublibre.org/rss/completo"}
TRACKERS = ["http://tracker.openbittorrent.com:80/announce",
            "udp://tracker.openbittorrent.com:6969/announce",
            "udp://tracker.torrent.eu.org:451",
            "udp://open.demonii.com:1337",
            "udp://tracker.opentrackr.org:1337/announce",
            "udp://tracker.cyberia.is:6969/announce"
            ]


# Función auxiliar para obtener el directorio de ejecución del script
def directorio_ex():
    if hasattr(sys, "frozen") or hasattr(sys, "importers"):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.realpath(__file__))


class Acerca_de(QtWidgets.QDialog):
    def __init__(self):
        super(Acerca_de, self).__init__()

        about_ui_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "about.ui"))
        uic.loadUi(about_ui_path, self)
        self.LbVersion.setText('Versión ' + VERSION + ' (Bisky & volante)')
        self.show()


# Ventana de opciones
class Ventana_opciones(QtWidgets.QDialog):
    def __init__(self, trackers, trackers_defecto, rss):
        super(Ventana_opciones, self).__init__()

        opciones_ui_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "opciones.ui"))
        uic.loadUi(opciones_ui_path, self)

        if trackers_defecto:
            self.checkBoxTrackers.setChecked(True)
        else:
            self.checkBoxTrackers.setChecked(False)
        self.EditorRSSSem.setText(rss['RSS semanal'])
        self.EditorRSSMen.setText(rss['RSS mensual'])
        self.EditorRSSTot.setText(rss['RSS total'])
        i = 1
        for tracker in trackers:
            editor = eval('self.EditorTracker' + str(i))
            editor.setText(tracker)
            i += 1
            if i > 6:
                break
        self.show()

    @pyqtSlot()
    def on_BotonCargarRSS_clicked(self):
        self.EditorRSSSem.setText(RSS_links['RSS semanal'])
        self.EditorRSSMen.setText(RSS_links['RSS mensual'])
        self.EditorRSSTot.setText(RSS_links['RSS total'])

    @pyqtSlot()
    def on_BotonCargarTrackers_clicked(self):
        for i in range(1, 7):
            editor = eval('self.EditorTracker' + str(i))
            editor.setText('')
        i = 1
        for tracker in TRACKERS:
            editor = eval('self.EditorTracker' + str(i))
            editor.setText(tracker)
            i += 1
            if i > 6:
                break


class DownloadThread(QThread):
    """Hilo para aislar el proceso de descarga,
    especialmente necesario en el caso del rss total"""

    def __init__(self, link, parent):
        """
        link: link de descarga
        fsize: tamaño total del archivo para """
        # super(DownloadThread, self).__init__(parent)
        QThread.__init__(self)
        self.link = link
        self.parent = parent
        self.start()

    def run(self):
        try:
            req = request.urlretrieve(self.link, "epublibre.rss", self.reporte)
            self.parent.fl = os.path.join(directorio_ex(), "epublibre.rss")
            self.parent.estado_descarga = True
        except Exception as e:
            print(e)
            self.parent.estado_descarga = False

    def reporte(self, blocknum, blocksize, totalsize):
        percent = round(blocknum / (blocksize / totalsize))
        self.parent.BarraBloque.setValue(percent)


class Servidor(QtWidgets.QMainWindow):
    def __init__(self):
        super(Servidor, self).__init__()
        self.RevisarTitulos = None
        self.lista_libros = []
        self.lista_archivos = []
        self.procesando = False
        self.filtrado = False
        self.libro_procesado = 0
        self.procesado_parcial = 0
        self.temporizador_on = False
        self.cancelar_borrado = False
        self.fl = ''
        self.cliente = TORRENT_CLIENTE
        self.trackers = TRACKERS
        self.trackers_defecto = True
        self.trackers_fichero = []
        self.rss_links = RSS_links
        self.dir_biblio = ''
        self.dir_duplicados = ''
        servidor_ui_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'servidor.ui'))
        uic.loadUi(servidor_ui_path, self)
        self.initUI()

    def initUI(self):
        self.SpinNumEnlaces.setValue(NUM_ENLACES)
        self.SpinVelEnvio.setValue(VELOCIDAD)
        self.BotonCancelar.setVisible(False)
        self.OpcionesServidor.setEnabled(False)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.lanza_enlace)
        for key in self.rss_links:
            if self.rss_links[key] != '':
                self.ComboRss.addItem(key)
        # Carga la configuración y estado del programa desde el archivo rss_server.ini
        self.show()
        QtWidgets.QApplication.processEvents()
        self.cargar_estado('rss_server.ini')

    @pyqtSlot()
    def on_BotonComenzar_clicked(self):
        self.comenzar_proceso()

    @pyqtSlot()
    def on_BotonPausa_clicked(self):
        self.pausar_proceso()

    @pyqtSlot()
    def on_BotonReiniciar_clicked(self):
        self.reiniciar_proceso()

    @pyqtSlot()
    def on_BotonDownload_clicked(self):
        # Descarga del fichero RSS
        self.muestra_mensaje('Descargando archivo %s de epublibre...' % self.ComboRss.currentText())
        # Habilitar / Deshabilitar botones y menús
        self.BotonPausa.setText("Pausa")
        self.OpcionesArchivos.setDisabled(True)
        self.BotonComenzar.setDisabled(True)
        self.BotonPausa.setDisabled(True)
        self.BotonReiniciar.setDisabled(True)
        self.OpcionesServidor.setDisabled(True)
        self.ContPestanas.setTabEnabled(1, False)
        self.menubar.setDisabled(True)
        # Doy el número de pasos a la barra parcial
        self.BarraBloque.setMaximum(100)
        self.BarraTotal.setValue(0)
        self.BarraBloque.setValue(0)
        LINK = self.rss_links[str(self.ComboRss.currentText())]
        try:
            self.estado_descarga = False
            descarga = DownloadThread(LINK, self)
            descarga.wait()
            if self.estado_descarga:
                self.BarraBloque.setValue(100)
                self.muestra_mensaje('Archivo RSS descargado correctamente', 'darkgreen')
                self.abre_fichero(False)
            else:
                self.muestra_mensaje('Error en la descarga del archivo del RSS', 'red')
                # Habilitar / Deshabilitar botones y menús
                self.OpcionesArchivos.setEnabled(True)
                self.BotonComenzar.setEnabled(True)
                self.BotonPausa.setEnabled(True)
                self.BotonReiniciar.setEnabled(True)
                self.OpcionesServidor.setEnabled(True)
                self.ContPestanas.setTabEnabled(1, True)
                self.menubar.setEnabled(True)
        except Exception as e:
            self.muestra_mensaje('Error en la descarga del archivo del RSS', 'red')
            self.muestra_mensaje(e, 'red')
            # Habilitar / Deshabilitar botones y menús
            self.OpcionesArchivos.setEnabled(True)
            self.BotonComenzar.setEnabled(True)
            self.BotonPausa.setEnabled(True)
            self.BotonReiniciar.setEnabled(True)
            self.OpcionesServidor.setEnabled(True)
            self.ContPestanas.setTabEnabled(1, True)
            self.menubar.setEnabled(True)

    @pyqtSlot()
    def on_BotonExaminar_clicked(self):
        # Apertura del fichero RSS o CSV
        ruta = os.path.dirname(self.fl)
        self.fl = QtWidgets.QFileDialog.getOpenFileName(self, 'Seleccionar archivo RSS o CSV', ruta,
                                                        'Archivos RSS o CSV (*.rss *.csv)')[0]
        if self.fl:
            self.fl = os.path.normpath(self.fl)
            self.abre_fichero()

    @pyqtSlot()
    def on_BotonCliente_clicked(self):
        dir_apertura = os.path.dirname(self.cliente)
        if sys.platform == "win32":
            patron = 'Archivos ejecutables (*.exe *.bat)'
            if dir_apertura == '':
                try:
                    dir_apertura = os.environ["PROGRAMFILES(X86)"]
                except:
                    dir_apertura = os.environ["PROGRAMFILES"]
        else:
            patron = 'Todos los archivos (*)'
            if dir_apertura == '':
                if os.path.exists("/usr/bin"):
                    dir_apertura = "/usr/bin"
                elif os.path.exists("/usr/local/bin"):
                    dir_apertura = "/usr/local/bin"
        cliente = QtWidgets.QFileDialog.getOpenFileName(self, 'Seleccionar cliente torrent', dir_apertura, patron)[0]
        if cliente:
            self.cliente = os.path.normpath(cliente)
            self.BotonCliente.setText(self.cliente)
            if 'utorrent' in os.path.normcase(cliente):
                self.EditDirDestino.setEnabled(True)
                self.BotonDirDestino.setEnabled(True)
            else:
                self.EditDirDestino.setDisabled(True)
                self.BotonDirDestino.setDisabled(True)

    @pyqtSlot()
    def on_BotonRutaDirectorio_clicked(self):
        # Selecciona el directorio de la biblioteca EPL
        ruta = self.dir_biblio
        dir_bib = QtWidgets.QFileDialog.getExistingDirectory(self, 'Seleccionar directorio', ruta,
                                                             QtWidgets.QFileDialog.ShowDirsOnly | QtWidgets.QFileDialog.DontResolveSymlinks | QtWidgets.QFileDialog.HideNameFilterDetails | QtWidgets.QFileDialog.ReadOnly)
        if dir_bib:
            self.dir_biblio = os.path.normpath(dir_bib)
            self.BotonRutaDirectorio.setText(os.path.normpath(dir_bib))
            self.BotonRutaBiblioteca.setText(os.path.normpath(dir_bib))
            self.OpcionesBorrado.setEnabled(True)
            self.OpcionesAccion.setEnabled(True)
            self.BotonBorrado.setEnabled(True)
            self.CompBiblio.setEnabled(True)

    @pyqtSlot()
    def on_BotonBorrado_clicked(self):
        # Realiza el borrado de los archivos duplicados
        self.ContPestanas.setTabEnabled(0, False)
        self.BotonBorrado.setVisible(False)
        self.BotonCancelar.setVisible(True)
        self.OpcionesBorrado.setDisabled(True)
        self.OpcionesAccion.setDisabled(True)
        self.OpcionesRuta.setDisabled(True)
        self.menubar.setDisabled(True)
        self.AreaMensajes.clear()
        self.cancelar_borrado = False
        exito = self.leer_directorio()
        if exito:
            self.borrar_antiguos()
        self.ContPestanas.setTabEnabled(0, True)
        self.BotonCancelar.setVisible(False)
        self.BotonBorrado.setVisible(True)
        self.OpcionesBorrado.setEnabled(True)
        self.OpcionesAccion.setEnabled(True)
        self.OpcionesRuta.setEnabled(True)
        self.menubar.setEnabled(True)

    @pyqtSlot()
    def on_BotonCancelar_clicked(self):
        self.cancelar_borrado = True

    @pyqtSlot()
    def on_BotonDirDestino_clicked(self):
        # Selecciona el directorio de destino de los enlaces
        ruta = self.EditDirDestino.text()
        if ruta == '':
            ruta = self.dir_biblio
        dir_bib = QtWidgets.QFileDialog.getExistingDirectory(self, 'Seleccionar directorio', ruta,
                                                             QtWidgets.QFileDialog.ShowDirsOnly | QtWidgets.QFileDialog.DontResolveSymlinks | QtWidgets.QFileDialog.HideNameFilterDetails | QtWidgets.QFileDialog.ReadOnly)
        if dir_bib:
            self.EditDirDestino.setText(os.path.normpath(dir_bib))

    @pyqtSlot()
    def on_actionGuardarEstado_triggered(self):
        self.guardar_estado("estado.sta")

    @pyqtSlot()
    def on_actionCargarEstado_triggered(self):
        self.cargar_estado("estado.sta")

    @pyqtSlot()
    def on_actionAcercaDe_triggered(self):
        acerca = Acerca_de()
        acerca.exec_()

    @pyqtSlot()
    def on_actionModificarOpciones_triggered(self):
        opciones = Ventana_opciones(self.trackers, self.trackers_defecto, self.rss_links)
        if opciones.exec_():
            # Guardado de las opciones modificadas
            self.rss_links['RSS semanal'] = opciones.EditorRSSSem.text()
            self.rss_links['RSS mensual'] = opciones.EditorRSSMen.text()
            self.rss_links['RSS total'] = opciones.EditorRSSTot.text()
            # Asignación de los nuevos RSS
            self.ComboRss.clear()
            for key in self.rss_links:
                if self.rss_links[key] != '':
                    self.ComboRss.addItem(key)
            # Asignación de los trackers
            self.trackers_defecto = opciones.checkBoxTrackers.isChecked()
            i = 1
            self.trackers = []
            while i < 7:
                editor = eval('opciones.EditorTracker' + str(i))
                tracker = editor.text().strip()
                if tracker != '':
                    self.trackers.append(tracker)
                i += 1

    @pyqtSlot()
    def on_actionReiniciarAjustes_triggered(self):
        # Establece todos los ajustes en sus valores por defecto
        elegido = QtWidgets.QMessageBox.question(None, "Reiniciar todos los ajustes del programa",
                                                 "¿Estás seguro de que quieres reiniciar todos los ajustes del programa a sus valores iniciales?",
                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                 QtWidgets.QMessageBox.No)
        if elegido == QtWidgets.QMessageBox.No:
            return
        self.lista_libros = []
        self.lista_archivos = []
        self.procesando = False
        self.libro_procesado = 0
        self.filtrado = False
        self.procesado_parcial = 0
        self.temporizador_on = False
        self.cancelar_borrado = False
        self.fl = ''
        self.cliente = TORRENT_CLIENTE
        self.trackers = TRACKERS
        self.trackers_defecto = True
        self.trackers_fichero = []
        self.rss_links = RSS_links
        self.dir_biblio = ''
        self.dir_duplicados = ''
        self.SpinNumEnlaces.setValue(NUM_ENLACES)
        self.SpinVelEnvio.setValue(VELOCIDAD)
        self.CompBiblio.setChecked(False)
        self.CompBiblio.setDisabled(True)
        self.BotonCancelar.setVisible(False)
        self.OpcionesServidor.setEnabled(False)
        self.ComboRss.clear()
        for key in self.rss_links:
            if self.rss_links[key] != '':
                self.ComboRss.addItem(key)
        self.BotonExaminar.setText('Examinar...')
        self.BotonCliente.setText('Seleccionar...')
        self.BotonRutaDirectorio.setText('Examinar...')
        self.BotonRutaBiblioteca.setText('Examinar...')
        self.EditDirDestino.setText('')
        self.EditDirDestino.setDisabled(True)
        self.BotonDirDestino.setDisabled(True)
        self.BarraTotal.setMaximum(100)
        self.BarraBloque.setMaximum(100)
        self.BarraTotal.setValue(0)
        self.BarraBloque.setValue(0)
        self.radioMover.setChecked(False)
        self.radioEliminar.setChecked(True)
        self.radioMover.setToolTip('Mover la versión más antigua del fichero a otro directorio (especificar)')
        # Habilitar/deshabilitar botones
        self.BotonComenzar.setDisabled(True)
        self.BotonPausa.setDisabled(True)
        self.BotonReiniciar.setDisabled(True)
        self.BotonPausa.setText("Pausa")
        self.actionComenzar.setDisabled(True)
        self.actionPausa.setDisabled(True)
        self.actionPausa.setText("Pausa")
        self.actionReiniciar.setDisabled(True)
        self.OpcionesServidor.setDisabled(True)
        self.CompBiblio.setDisabled(True)
        self.ContPestanas.setTabEnabled(1, True)
        self.OpcionesArchivos.setEnabled(True)
        self.OpcionesBorrado.setDisabled(True)
        self.OpcionesAccion.setDisabled(True)
        self.BotonBorrado.setDisabled(True)
        self.RevisarTitulos.setChecked(True)
        self.CoincSinId.setChecked(True)
        self.actionAbrir.setEnabled(True)
        self.actionCliente.setEnabled(True)
        self.actionGuardarEstado.setEnabled(True)
        self.actionCargarEstado.setEnabled(True)
        self.actionDirectorio.setEnabled(True)
        self.AreaMensajes.clear()
        self.muestra_mensaje('Ajustes del programa establecidos en sus valores iniciales', 'darkgreen')

    @pyqtSlot()
    def on_CompBiblio_clicked(self):
        if self.CompBiblio.isChecked():
            QtWidgets.QMessageBox.information(None, "Desactivación de guardado de estado",
                                              "Al activar esta opción no será posible guardar el estado de progreso del programa para continuar más adelante.\n" +
                                              "El motivo es que el contenido del directorio donde se guarda la biblioteca puede variar su contenido según se vayan descargando libros nuevos.",
                                              QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)
        elif self.filtrado:
            self.abre_fichero(False)
            self.filtrado = False

    @pyqtSlot()
    def on_radioMover_clicked(self):
        if self.radioMover.isChecked():
            # Selecciona el directorio de la biblioteca EPL
            ruta = self.dir_duplicados
            dir_bib = QtWidgets.QFileDialog.getExistingDirectory(self, 'Seleccionar directorio para los duplicados',
                                                                 ruta,
                                                                 QtWidgets.QFileDialog.ShowDirsOnly | QtWidgets.QFileDialog.DontResolveSymlinks | QtWidgets.QFileDialog.HideNameFilterDetails | QtWidgets.QFileDialog.ReadOnly)
            if dir_bib:
                if os.path.normpath(dir_bib) != self.dir_biblio:
                    self.dir_duplicados = os.path.normpath(dir_bib)
                    self.radioMover.setToolTip(
                        'Mover la versión más antigua del fichero al directorio ' + self.dir_duplicados)
                else:
                    QtWidgets.QMessageBox.warning(None, "Error de directorio",
                                                  "No se puede especificar como directorio de destino de los duplicados\n" +
                                                  "el mismo directorio donde está ubicada la biblioteca.",
                                                  QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)
                    self.radioEliminar.setChecked(True)
            else:
                self.radioEliminar.setChecked(True)

    @pyqtSlot()
    def on_radioEliminar_clicked(self):
        if self.radioEliminar.isChecked():
            # Cambio el tooltip de la otra opción
            self.radioMover.setToolTip('Mover la versión más antigua del fichero a otro directorio (especificar)')

    def closeEvent(self, event):
        # Cancelo posibles procesos en marcha
        self.cancelar_borrado = True
        # Guarda la configuración y estado del programa en el archivo rss_server.ini
        self.guardar_estado('rss_server.ini')

    def abre_fichero(self, borrar_mensajes=True):
        if os.path.isfile(self.fl):
            try:
                QtWidgets.QApplication.setOverrideCursor(QtGui.QCursor(Qt.WaitCursor))
                # Habilitar / Deshabilitar botones y menús
                self.BotonPausa.setText("Pausa")
                self.OpcionesArchivos.setDisabled(True)
                self.BotonComenzar.setDisabled(True)
                self.BotonPausa.setDisabled(True)
                self.BotonReiniciar.setDisabled(True)
                self.OpcionesServidor.setDisabled(True)
                self.ContPestanas.setTabEnabled(1, False)
                self.menubar.setDisabled(True)
                # Borrar área de mensajes e indicar la carga del archivo
                if borrar_mensajes:
                    self.AreaMensajes.clear()
                self.muestra_mensaje('Procesando archivo ' + self.fl + '...')
                QtWidgets.QApplication.processEvents()
                # Procesar el archivo, dependiendo de su tipo
                nombre_archivo, extension = os.path.splitext(self.fl)
                if extension.lower() == '.rss':
                    exito = self.readFile_XML(self.fl)
                else:
                    if extension.lower() == '.csv':
                        exito = self.readFile_CSV(self.fl)
                    else:
                        self.muestra_mensaje('El tipo de archivo cargado es incorrecto (sólo se admiten RSS y CSV)',
                                             'red')
                        exito = False
                if exito:
                    # Reinicializar variables
                    self.libro_procesado = 0
                    self.procesado_parcial = 0
                    # Habilitar el botón de comenzar y los spinboxes
                    self.BotonComenzar.setEnabled(True)
                    self.actionComenzar.setEnabled(True)
                    self.OpcionesServidor.setEnabled(True)
                    self.BotonExaminar.setText(self.fl)
            finally:
                self.OpcionesArchivos.setEnabled(True)
                self.actionAbrir.setEnabled(True)
                self.actionCliente.setEnabled(True)
                self.ContPestanas.setTabEnabled(1, True)
                self.menubar.setEnabled(True)
                QtWidgets.QApplication.restoreOverrideCursor()
        else:
            self.muestra_mensaje('No se pudo encontrar el archivo indicado: ' + str(self.fl), 'red')

    def muestra_mensaje(self, mensaje, color="black"):
        # Mostrar mensaje en área de notificación en el color indicado
        self.AreaMensajes.setTextColor(QtGui.QColor(color))
        self.AreaMensajes.append(mensaje)
        return

    def lanza_enlace(self):
        # Enviar enlace al programa torrent
        if not self.temporizador_on:
            self.temporizador_on = True
            if self.libro_procesado < len(self.lista_libros):
                libro = self.lista_libros[self.libro_procesado]
                if libro:
                    try:
                        if "enlace" in libro:
                            enlace = 'magnet:?xt=urn:btih:' + libro["enlace"] + '&dn=EPL_[' + libro["epl_id"] + ']_'
                            enlace += str(normalize('NFKD', str(libro["titulo"])).encode('ASCII', 'ignore'),
                                          encoding='UTF-8')
                            enlace += self.trackers_usar
                            if sys.platform == "win32":
                                if self.cliente != '':
                                    if (self.EditDirDestino.text() != '') and ('utorrent' in self.cliente.lower()):
                                        subprocess.Popen(
                                            [self.cliente, '/DIRECTORY', os.path.normpath(self.EditDirDestino.text()),
                                             enlace])
                                    else:
                                        subprocess.Popen([self.cliente, enlace])
                                else:
                                    os.startfile(enlace)
                            else:
                                try:
                                    if (self.cliente != '') and (os.path.isfile(self.cliente)):
                                        if ((self.EditDirDestino.text() != '') and (
                                                os.path.isdir(self.EditDirDestino.text()))
                                                and ('utorrent' in os.path.normcase(self.cliente))):
                                            subprocess.Popen(
                                                [self.cliente, '/DIRECTORY', self.EditDirDestino.text(), enlace])
                                        else:
                                            subprocess.Popen([self.cliente, enlace])
                                    elif sys.platform == "darwin" or sys.platform == "mac":
                                        subprocess.Popen(['open', enlace])
                                    else:
                                        try:
                                            subprocess.Popen(['xdg-open', enlace])
                                        except:
                                            try:
                                                subprocess.Popen(['gnome-open', enlace])
                                            except:
                                                subprocess.Popen(['mate-open', enlace])
                                except:
                                    self.muestra_mensaje(
                                        'No fue posible usar el programa asociado a los archivos torrent',
                                        'red')
                                    self.pausar_proceso()
                                    return
                        if "titulo" in libro:
                            info = '(' + str(self.libro_procesado + 1) + '/' + str(len(self.lista_libros)) + ')'
                            self.muestra_mensaje(info + ' Enviando enlace del libro ' + libro["titulo"] + '...', 'blue')
                    except Exception as ex:
                        self.muestra_mensaje('Se ha producido un error al enviar el enlace', 'red')
                        self.muestra_mensaje(str(ex))
                    self.libro_procesado += 1
                    self.procesado_parcial += 1
                    if self.procesado_parcial <= self.BarraBloque.maximum():
                        self.BarraBloque.setValue(self.procesado_parcial)
                    if self.libro_procesado <= self.BarraTotal.maximum():
                        self.BarraTotal.setValue(self.libro_procesado)
                    if self.libro_procesado >= len(self.lista_libros):
                        self.reiniciar_proceso(fin_normal=True)
                    elif self.procesado_parcial >= self.num_enlaces:
                        self.pausar_proceso()
                        self.procesado_parcial = 0
            else:
                self.reiniciar_proceso(fin_normal=True)
            self.temporizador_on = False

    def comenzar_proceso(self):
        # Comenzar envío de enlaces
        if not self.procesando:
            self.procesando = True
            # Comprobación de que los spinboxes tengan datos y sean correctos
            veloc = self.validar_spins(self.SpinVelEnvio.value(), self.SpinNumEnlaces.value())
            intervalo = round(1000 / veloc)
            # Habilitar / Deshabilitar botones
            self.BotonComenzar.setDisabled(True)
            self.BotonPausa.setDisabled(True)
            self.BotonReiniciar.setDisabled(True)
            self.OpcionesArchivos.setDisabled(True)
            self.actionCliente.setDisabled(True)
            self.actionComenzar.setDisabled(True)
            self.actionPausa.setDisabled(True)
            self.actionReiniciar.setDisabled(True)
            self.actionAbrir.setDisabled(True)
            self.actionGuardarEstado.setDisabled(True)
            self.actionCargarEstado.setDisabled(True)
            self.OpcionesServidor.setDisabled(True)
            self.ContPestanas.setTabEnabled(1, False)
            self.menubar.setDisabled(True)
            self.BarraTotal.setValue(0)
            # Si está marcada la opción de comprobar con la biblioteca descargada, realiza las comprobaciones pertinentes
            if self.CompBiblio.isChecked():
                self.filtrar_libros()
            # Doy el número de pasos a las barras total y parcial
            self.BarraTotal.setMaximum(len(self.lista_libros))
            if len(self.lista_libros) < self.num_enlaces:
                self.BarraBloque.setMaximum(len(self.lista_libros))
            else:
                self.BarraBloque.setMaximum(self.num_enlaces)
            self.BarraBloque.setValue(0)
            self.BarraTotal.setValue(0)
            # Reinicializar variables
            self.libro_procesado = 0
            self.procesado_parcial = 0
            # Asignación de trackers
            self.trackers_usar = ''
            if self.trackers_defecto and len(self.trackers_fichero) > 0:
                for tracker in self.trackers_fichero:
                    self.trackers_usar += '&tr=' + tracker
            elif len(self.trackers) > 0:
                for tracker in self.trackers:
                    self.trackers_usar += '&tr=' + tracker
            else:
                for tracker in TRACKERS:
                    self.trackers_usar += '&tr=' + tracker
            # Activación del temporizador
            self.timer.start(intervalo)
            # Indicar el inicio del proceso
            self.muestra_mensaje('Iniciando envío de enlaces al programa torrent...', 'darkgreen')
        # Habilitar / Deshabilitar botones
        self.BotonPausa.setEnabled(True)
        self.actionPausa.setEnabled(True)

    def pausar_proceso(self):
        # Detener envío de enlaces
        if self.procesando:
            self.timer.stop()
            self.procesando = False
            self.muestra_mensaje(
                'Envío de enlaces al programa torrent en pausa (pulsar el botón "Continuar" para reanudar)',
                'darkgreen')
            # Habilitar / Deshabilitar botones
            self.BotonPausa.setText("Continuar")
            self.BotonReiniciar.setEnabled(True)
            self.OpcionesArchivos.setEnabled(True)
            self.CompBiblio.setDisabled(True)
            self.actionPausa.setText("Continuar")
            self.actionAbrir.setEnabled(True)
            self.actionGuardarEstado.setEnabled(True)
            self.actionCargarEstado.setEnabled(True)
            self.actionCliente.setEnabled(True)
            self.OpcionesServidor.setEnabled(True)
            self.menubar.setEnabled(True)
        else:
            self.procesando = True
            # Comprobación de que los spinboxes tengan datos y sean correctos
            veloc = self.validar_spins(self.SpinVelEnvio.value(), self.SpinNumEnlaces.value())
            intervalo = round(1000 / veloc)
            # Ajuste del tamaño de la barra de bloque
            if (len(self.lista_libros) - self.libro_procesado) < self.num_enlaces:
                self.BarraBloque.setMaximum(len(self.lista_libros) - self.libro_procesado)
            else:
                self.BarraBloque.setMaximum(self.num_enlaces)
            # Asignación de trackers
            self.trackers_usar = ''
            if self.trackers_defecto and len(self.trackers_fichero) > 0:
                for tracker in self.trackers_fichero:
                    self.trackers_usar += '&tr=' + tracker
            elif len(self.trackers) > 0:
                for tracker in self.trackers:
                    self.trackers_usar += '&tr=' + tracker
            else:
                for tracker in TRACKERS:
                    self.trackers_usar += '&tr=' + tracker
            # Activación del temporizador
            self.timer.start(intervalo)
            self.muestra_mensaje('Envío de enlaces al programa torrent reanudado', 'darkgreen')
            # Habilitar / Deshabilitar botones
            self.BotonPausa.setText("Pausa")
            self.BotonReiniciar.setDisabled(True)
            self.OpcionesArchivos.setDisabled(True)
            self.actionPausa.setText("Pausa")
            self.actionAbrir.setDisabled(True)
            self.actionGuardarEstado.setDisabled(True)
            self.actionCargarEstado.setDisabled(True)
            self.actionCliente.setDisabled(True)
            self.OpcionesServidor.setDisabled(True)
            self.menubar.setDisabled(True)
        self.BotonComenzar.setDisabled(True)

    def reiniciar_proceso(self, fin_normal=False):
        # Confirmar la detención
        if not fin_normal:
            elegido = QtWidgets.QMessageBox.question(None, "Reiniciar proceso",
                                                     "¿Estás seguro de que quieres reiniciar el proceso? (el estado actual se perderá)",
                                                     QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                     QtWidgets.QMessageBox.No)
            if elegido == QtWidgets.QMessageBox.No:
                return
        # Detener y reiniciar envío de enlaces
        if self.procesando:
            self.timer.stop()
            self.procesando = False
        self.libro_procesado = 0
        self.procesado_parcial = 0
        if fin_normal:
            self.muestra_mensaje('Envío de enlaces al programa torrent finalizado', 'darkgreen')
        else:
            # Borrar área de mensajes e indicar el final del proceso
            self.AreaMensajes.clear()
            self.muestra_mensaje(
                'Envío de enlaces al programa torrent reiniciado (pulsar el botón "Comenzar" para empezar de nuevo)',
                'darkgreen')
            # Reinicializar barras de progreso
            self.BarraBloque.setValue(0)
            self.BarraTotal.setValue(0)
        # Habilitar / Deshabilitar botones
        self.BotonComenzar.setEnabled(True)
        self.BotonPausa.setDisabled(True)
        self.BotonReiniciar.setDisabled(True)
        self.BotonPausa.setText("Pausa")
        self.OpcionesArchivos.setEnabled(True)
        self.CompBiblio.setEnabled(True)
        self.actionCliente.setEnabled(True)
        self.actionComenzar.setEnabled(True)
        self.actionPausa.setDisabled(True)
        self.actionPausa.setText("Pausa")
        self.actionReiniciar.setDisabled(True)
        self.actionAbrir.setEnabled(True)
        self.actionGuardarEstado.setEnabled(True)
        self.actionCargarEstado.setEnabled(True)
        self.OpcionesServidor.setEnabled(True)
        self.ContPestanas.setTabEnabled(1, True)
        self.menubar.setEnabled(True)

    def guardar_estado(self, archivo="estado.sta"):
        # Guarda el estado actual del programa en el archivo "estado.sta" o en que se haya pasado
        datos = {}
        datos["archivo"] = self.fl
        datos["cliente"] = self.cliente
        datos["num_enlaces"] = self.SpinNumEnlaces.value()
        datos["velocidad"] = self.SpinVelEnvio.value()
        if not self.CompBiblio.isChecked():
            datos["proc_bloque"] = self.procesado_parcial
            datos["proc_total"] = self.libro_procesado
        else:
            datos["proc_bloque"] = datos["proc_total"] = 0
        datos["dir_destino"] = self.EditDirDestino.text()
        datos["dir_biblio"] = self.dir_biblio
        datos["dir_duplicados"] = self.dir_duplicados
        datos["revisar_titulos"] = int(self.RevisarTitulos.isChecked())
        datos["coincidir_sin_id"] = int(self.CoincSinId.isChecked())
        datos["comprobar_biblio"] = int(self.CompBiblio.isChecked())
        datos["accion_duplicados"] = int(self.radioEliminar.isChecked())
        datos["trackers"] = self.trackers
        datos["trackers_defecto"] = int(self.trackers_defecto)
        datos["rss_links"] = self.rss_links
        try:
            arch = os.path.join(directorio_ex(), archivo)
            with open(arch, 'w') as f:
                json.dump(datos, f, ensure_ascii=False)
            self.muestra_mensaje('Estado actual del programa guardado con éxito', 'darkgreen')
        except Exception as ex:
            self.muestra_mensaje('Se ha producido un error al intentar guardar el estado del programa en el disco:',
                                 'red')
            self.muestra_mensaje(str(ex), 'red')

    def cargar_estado(self, archivo="estado.sta"):
        # Carga el estado anteriormente guardado en el archivo "estado.sta" o en el que se le pase
        # Confirmar la carga
        if archivo == "estado.sta":
            elegido = QtWidgets.QMessageBox.question(None, "Cargar estado anterior",
                                                     "¿Estás seguro de que quieres cargar el estado anterior? (el estado actual se perderá)",
                                                     QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                     QtWidgets.QMessageBox.No)
            if elegido == QtWidgets.QMessageBox.No:
                return
        # Comprobación de existencia del fichero
        arch = os.path.join(directorio_ex(), archivo)
        if not os.path.isfile(arch):
            if archivo == "estado.sta":
                self.muestra_mensaje('No hay ningún estado anterior guardado', 'red')
            return
        # Carga de datos
        try:
            try:
                with open(arch, 'r') as f:
                    datos = json.load(f)
            except Exception as ex:
                self.muestra_mensaje(
                    'Se ha producido un error al intentar cargar el estado anterior del programa desde el disco:',
                    'red')
                self.muestra_mensaje(str(ex), 'red')
                return
            # Procesado de los datos
            self.AreaMensajes.clear()
            self.fl = datos["archivo"]
            if self.fl != '':
                self.abre_fichero()
            else:
                self.lista_libros = {}
                self.BotonExaminar.setText('Examinar...')
            self.cliente = datos["cliente"]
            if self.cliente != '':
                self.BotonCliente.setText(os.path.normpath(self.cliente))
                if 'utorrent' in self.cliente.lower():
                    self.EditDirDestino.setEnabled(True)
                    self.BotonDirDestino.setEnabled(True)
                else:
                    self.EditDirDestino.setDisabled(True)
                    self.BotonDirDestino.setDisabled(True)
            else:
                self.BotonCliente.setText('Seleccionar...')
                self.EditDirDestino.setDisabled(True)
                self.BotonDirDestino.setDisabled(True)
            if 'dir_destino' in datos and datos["dir_destino"] != '':
                self.EditDirDestino.setText(os.path.normpath(datos["dir_destino"]))
            if 'trackers' in datos:
                self.trackers = datos["trackers"]
                if len(self.trackers) == 0:
                    self.trackers = TRACKERS
            else:
                self.trackers = TRACKERS
            if 'trackers_defecto' in datos:
                self.trackers_defecto = bool(datos["trackers_defecto"])
            if 'rss_links' in datos:
                self.rss_links = datos["rss_links"]
                if len(self.rss_links) == 0:
                    self.rss_links = RSS_links
            else:
                self.rss_links = RSS_links

            # Asignación de los valores de RSS al combo
            self.ComboRss.clear()
            for key in self.rss_links:
                if self.rss_links[key] != '':
                    self.ComboRss.addItem(key)
            # Comprobación de que los valores de los valores de velocidad de número de enlaces cargados sean correctos
            veloc = self.validar_spins(datos["velocidad"], datos["num_enlaces"])
            intervalo = round(1000 / veloc)
            # Cargo los datos de procesado parcial y total
            try:
                self.procesado_parcial = int(datos["proc_bloque"])
                if (self.procesado_parcial > len(self.lista_libros)) or (self.procesado_parcial > self.num_enlaces):
                    self.procesado_parcial = 0
            except:
                self.procesado_parcial = 0
            try:
                self.libro_procesado = int(datos["proc_total"])
                if self.libro_procesado > len(self.lista_libros):
                    self.libro_procesado = 0
            except:
                self.libro_procesado = 0
            # Doy el número de pasos a las barras total y parcial
            if len(self.lista_libros) > 0:
                self.BarraTotal.setMaximum(len(self.lista_libros))
                if len(self.lista_libros) < self.num_enlaces:
                    self.BarraBloque.setMaximum(len(self.lista_libros))
                else:
                    self.BarraBloque.setMaximum(self.num_enlaces)
                self.BarraBloque.setValue(self.procesado_parcial)
                self.BarraTotal.setValue(self.libro_procesado)
            else:
                self.BarraTotal.setMaximum(100)
                self.BarraBloque.setMaximum(100)
                self.BarraTotal.setValue(0)
                self.BarraBloque.setValue(0)
            # Carga del checkbox de comprobar biblioteca
            self.CompBiblio.setChecked(bool(datos["comprobar_biblio"]))
            # Intervalo del temporizador
            self.timer.setInterval(intervalo)
            # Carga de valores de la pestaña de archivos
            self.dir_biblio = datos["dir_biblio"]
            if self.dir_biblio != '':
                self.BotonRutaDirectorio.setText(os.path.normpath(self.dir_biblio))
                self.BotonRutaBiblioteca.setText(os.path.normpath(self.dir_biblio))
                self.OpcionesBorrado.setEnabled(True)
                self.OpcionesAccion.setEnabled(True)
                self.BotonBorrado.setEnabled(True)
                self.CompBiblio.setEnabled(True)
            else:
                self.BotonRutaDirectorio.setText('Examinar...')
                self.BotonRutaBiblioteca.setText('Examinar...')
                self.OpcionesBorrado.setDisabled(True)
                self.OpcionesAccion.setDisabled(True)
                self.BotonBorrado.setDisabled(True)
                self.CompBiblio.setDisabled(True)
            if datos["dir_biblio"] != datos["dir_duplicados"]:
                self.dir_duplicados = datos["dir_duplicados"]
            self.RevisarTitulos.setChecked(bool(datos["revisar_titulos"]))
            self.CoincSinId.setChecked(bool(datos["coincidir_sin_id"]))
            if datos["accion_duplicados"] == 1:
                self.radioEliminar.setChecked(True)
            else:
                self.radioMover.setChecked(True)
                self.radioMover.setToolTip(
                    'Mover la versión más antigua del fichero al directorio ' + self.dir_duplicados)
            # Habilitar/deshabilitar botones
            if len(self.lista_libros) == 0:
                self.BotonComenzar.setDisabled(True)
                self.BotonPausa.setDisabled(True)
                self.BotonReiniciar.setDisabled(True)
                self.BotonPausa.setText("Pausa")
                self.actionComenzar.setDisabled(True)
                self.actionPausa.setDisabled(True)
                self.actionPausa.setText("Pausa")
                self.actionReiniciar.setDisabled(True)
                self.OpcionesServidor.setDisabled(True)
                self.ContPestanas.setTabEnabled(1, True)
            else:
                self.OpcionesServidor.setEnabled(True)
                if self.libro_procesado == 0:
                    self.BotonComenzar.setEnabled(True)
                    self.BotonPausa.setDisabled(True)
                    self.BotonReiniciar.setDisabled(True)
                    self.BotonPausa.setText("Pausa")
                    self.actionComenzar.setEnabled(True)
                    self.actionPausa.setDisabled(True)
                    self.actionPausa.setText("Pausa")
                    self.actionReiniciar.setDisabled(True)
                    self.ContPestanas.setTabEnabled(1, True)
                else:
                    self.BotonComenzar.setDisabled(True)
                    self.BotonPausa.setEnabled(True)
                    self.BotonReiniciar.setEnabled(True)
                    self.BotonPausa.setText("Continuar")
                    self.actionComenzar.setDisabled(True)
                    self.actionPausa.setEnabled(True)
                    self.actionPausa.setText("Continuar")
                    self.actionReiniciar.setEnabled(True)
                    self.CompBiblio.setDisabled(True)
                    self.ContPestanas.setTabEnabled(1, False)
            self.OpcionesArchivos.setEnabled(True)
            self.actionAbrir.setEnabled(True)
            self.actionCliente.setEnabled(True)
            self.actionGuardarEstado.setEnabled(True)
            self.actionCargarEstado.setEnabled(True)
            self.muestra_mensaje('Estado anterior del programa cargado con éxito', 'darkgreen')
        except:
            self.muestra_mensaje('Se ha producido un error al cargar el estado anterior del programa', 'red')

    def validar_spins(self, velocidad, num_enlaces):
        # Comprobación de que los valores de los valores de velocidad de número de enlaces cargados sean correctos
        try:
            veloc = int(velocidad)
            if veloc < 1 or veloc > 99999:
                veloc = VELOCIDAD
        except:
            veloc = VELOCIDAD
        self.SpinVelEnvio.setValue(veloc)
        try:
            self.num_enlaces = int(num_enlaces)
            if self.num_enlaces < 1 or self.num_enlaces > 99999:
                self.num_enlaces = NUM_ENLACES
        except:
            self.num_enlaces = NUM_ENLACES
        self.SpinNumEnlaces.setValue(self.num_enlaces)
        return veloc

    def readFile_XML(self, filename):
        exito = False
        self.lista_libros = []
        filename = os.path.normpath(filename)
        if not os.path.isfile(filename):
            self.muestra_mensaje('No se pudo acceder al fichero indicado', 'red')
            return exito
        try:
            # Parseo del contenido del archivo comprimido
            midom = xml.dom.minidom.parse(filename)
            # Localización del nodo principal del canal
            canal = midom.getElementsByTagName("channel").item(0)
            if canal is not None:
                # Comprobación de que se trata del RSS de EPL
                titulo = canal.getElementsByTagName("title").item(0).firstChild.data
                if titulo != 'RSS de enlaces Epublibre':
                    self.muestra_mensaje('El archivo indicado no es el RSS de Epublibre', 'red')
                    midom.unlink()
                    return exito
                # Comprobación de que no se trata de un archivo de error por haber sobrepasado el límite de descargas del fichero
                descripcion = canal.getElementsByTagName("description").item(0).firstChild.data
                if 'Se ha superado el número de descargas diarias' in descripcion:
                    self.muestra_mensaje(
                        'El archivo descargado no es válido por haber superado el límite de descargas diario:', 'red')
                    self.muestra_mensaje(descripcion, 'red')
                    midom.unlink()
                    return exito
                # Búsqueda de los items
                items_encontrados = canal.getElementsByTagName("item")
                total = len(items_encontrados)
                # Lectura de los trackers del primer enlace (se supone que son los mismos para todos ellos)
                item = items_encontrados[0]
                nodo_atributo = item.getElementsByTagName("link").item(0)
                if nodo_atributo is not None:
                    enlace = nodo_atributo.firstChild.wholeText
                    self.trackers_fichero = re.findall("tr=(.*?)(?:&|$)", enlace)
                # Para cada item lectura del atributo title, y del enlace
                for item in items_encontrados:
                    QtWidgets.QApplication.processEvents()
                    datos_item = {}
                    # Lectura del título
                    nodo_atributo = item.getElementsByTagName("title").item(0)
                    if nodo_atributo is not None:
                        titulo = nodo_atributo.firstChild.wholeText
                        # Borrado de salto de línea del principio
                        titulo = re.sub("^\\n", '', titulo)
                        # Borrado de los géneros
                        titulo = re.sub("\[(.)*\]", '', titulo)
                        # Borrado de salto de línea y espacios del final
                        datos_item["titulo"] = self.normalizar(re.sub("\\n+( )+$", '', titulo))
                    # Lectura del autor
                    nodo_atributo = item.getElementsByTagName("autor").item(0)
                    if nodo_atributo is not None:
                        autor = nodo_atributo.firstChild.wholeText
                        # Borrado de salto de línea del principio
                        autor = re.sub("^\\n", '', autor)
                        # Borrado de salto de línea y espacios del final
                        datos_item["autor"] = self.normalizar(re.sub("\\n+( )+$", '', autor))
                    # Lectura de la versión
                    nodo_atributo = item.getElementsByTagName("rev").item(0)
                    if nodo_atributo is not None:
                        datos_item["version"] = nodo_atributo.firstChild.data
                    # Lectura del enlace y el epl_id
                    nodo_atributo = item.getElementsByTagName("link").item(0)
                    if nodo_atributo is not None:
                        enlace = nodo_atributo.firstChild.wholeText
                        # Borrado del salto de línea del principio
                        enlace = re.sub("^\\n", '', enlace)
                        # Borrado del salto de línea y espacios del final
                        enlace = re.sub("\\n+( )+$", '', enlace)
                        # Extracción de la cadena del HASH y el EPL_ID
                        enl = re.search("btih:(.*?)&dn=", enlace)
                        if enl:
                            datos_item["enlace"] = enl.group(1)
                        epl_id = re.search("&dn=EPL_\[(.*?)\]", enlace)
                        if epl_id:
                            datos_item["epl_id"] = epl_id.group(1)

                    # Adición del libro a la lista
                    self.lista_libros.append(datos_item)
            midom.unlink()
            self.muestra_mensaje('Archivo cargado con éxito (' + str(len(self.lista_libros)) + ' enlaces encontrados)',
                                 'darkgreen')
            exito = True
        except Exception as ex:
            self.muestra_mensaje('Se ha producido el siguiente error al leer el archivo RSS:', 'red')
            self.muestra_mensaje(str(ex), 'red')
            if 'midom' in locals():
                midom.unlink()
        return exito

    def readFile_CSV(self, filename):
        exito = False
        self.lista_libros = []
        filename = os.path.normpath(filename)
        if not os.path.isfile(filename):
            self.muestra_mensaje('No se pudo acceder al fichero indicado', 'red')
            return exito
        try:
            f = open(filename, "r", encoding="utf8")
        except:
            self.muestra_mensaje('No se pudo abrir el fichero indicado', 'red')
            return exito
        # Localización de las entradas del título y los enlaces
        try:
            linea = f.readline()
            # Borrado de comilla del principio
            linea = re.sub('^"', '', linea)
            # Borrado de salto de línea y espacios del final
            linea = re.sub('"\\n+( )*$', '', linea)
            encabezados = linea.split('","')
            try:
                ind_epl_id = encabezados.index("EPL Id")
                ind_titulo = encabezados.index("Título")
                ind_autor = encabezados.index("Autor")
                ind_version = encabezados.index("Revisión")
                try:
                    ind_enlace = encabezados.index("Enlace(s)")
                except:
                    ind_enlace = encabezados.index('Enlace(s)"')
            except:
                self.muestra_mensaje('El fichero indicado no es un CSV de Epublibre', 'red')
                return exito
            del encabezados
            # Lectura del resto del fichero
            while True:
                QtWidgets.QApplication.processEvents()
                linea = f.readline()
                if not linea:
                    break
                # Borrado de comilla del principio
                linea = re.sub('^"', '', linea)
                # Borrado de salto de línea y espacios del final
                linea = re.sub('"\\n+( )*$', '', linea)
                contenido = linea.split('","')
                enlaces = contenido[ind_enlace].split(', ')
                for enl in enlaces:
                    datos_item = {}
                    datos_item["epl_id"] = contenido[ind_epl_id]
                    datos_item["titulo"] = self.normalizar(contenido[ind_titulo])
                    autor = contenido[ind_autor]
                    if autor.count(" & ") > 1:
                        autor = 'AA. VV.'
                    datos_item["autor"] = self.normalizar(autor)
                    datos_item["version"] = contenido[ind_version]
                    datos_item["enlace"] = enl
                    # Adición del libro a la lista
                    self.lista_libros.append(datos_item)
            exito = True
            self.muestra_mensaje('Archivo cargado con éxito(' + str(len(self.lista_libros)) + ' enlaces encontrados)',
                                 'darkgreen')
        except Exception as ex:
            self.muestra_mensaje('Se ha producido el siguiente error al leer el archivo CSV:', 'red')
            self.muestra_mensaje(str(ex), 'red')
        f.close()
        return exito

    def leer_directorio(self, incluir_autor=False):
        # Procesa todo el contenido de un directorio y almacena en una lista el ID, versión y título de cada archivo
        self.muestra_mensaje('Revisando directorio ' + self.dir_biblio + '...', 'black')
        try:
            self.lista_archivos = []
            ficheros = os.listdir(self.dir_biblio)
            # Doy el número de pasos a las barras total y parcial
            if not incluir_autor:
                self.BarraTotal.setMaximum(2 * len(ficheros))
                self.BarraTotal.setValue(0)
            self.BarraBloque.setMaximum(len(ficheros))
            self.BarraBloque.setValue(0)
            for fichero in ficheros:
                QtWidgets.QApplication.processEvents()
                if os.path.isdir(fichero):
                    leer_directorio(fichero)
                else:
                    archivo = {}
                    nom, ext = os.path.splitext(fichero)
                    archivo['ext'] = ext.replace('.', '').strip()
                    # Extraigo la versión del nombre
                    nom = nom.strip()
                    coinc = re.search("\([rv]([0-9]*\.[0-9]*)(.*?)\)$", nom)
                    if coinc:
                        archivo['version'] = coinc.group(1)
                        nom = re.sub("\([rv]([0-9]*\.[0-9]*)(.*?)\)$", '', nom).strip()
                    # Extraigo el ID si existe
                    coinc = re.search("\[([0-9]*)\]$", nom)
                    if coinc:
                        archivo['ID'] = coinc.group(1)
                        nom = re.sub("\[([0-9]*)\]$", '', nom).strip()
                    archivo['titulo_comp'] = nom
                    archivo['nombre'] = fichero
                    if incluir_autor:
                        try:
                            # Elimino las colecciones e IDs
                            nom = re.sub("\[(.*?)\]", '', nom)
                            partes = nom.split(' - ', 1)
                            if len(partes) < 2:
                                partes = nom.split('- ', 1)
                                if len(partes) < 2:
                                    partes = nom.split(' -', 1)
                                    if len(partes) < 2:
                                        partes = nom.split('-', 1)
                                    if len(partes) < 2:
                                        continue
                            autor = partes[0].strip()
                            nom_autores = ''
                            autores = autor.split('&')
                            for aut in autores:
                                nom_autor = ''
                                p_autor = aut.split(',', 1)
                                for i in range(len(p_autor) - 1, -1, -1):
                                    if nom_autor != '':
                                        nom_autor += ' '
                                    nom_autor += p_autor[i].strip()
                                if nom_autores != '':
                                    nom_autores += ' & '
                                nom_autores += nom_autor
                            archivo['autor'] = self.normalizar(nom_autores)
                            archivo['titulo'] = self.normalizar(partes[1].strip())
                        except:
                            pass
                    self.lista_archivos.append(archivo)
                    # Actualizar barras de progreso
                    if self.BarraBloque.value() <= self.BarraBloque.maximum():
                        self.BarraBloque.setValue(self.BarraBloque.value() + 1)
                    if not incluir_autor:
                        if self.BarraTotal.value() <= self.BarraTotal.maximum():
                            self.BarraTotal.setValue(self.BarraTotal.value() + 1)
                # Comprobación de proceso cancelado
                if self.cancelar_borrado:
                    self.muestra_mensaje('Proceso de borrado cancelado', 'darkgreen')
                    return False
            self.muestra_mensaje('El directorio se ha revisado con éxito.', 'darkgreen')
            return True
        except Exception as ex:
            self.muestra_mensaje('Se ha producido el siguiente error al revisar el directorio:', 'red')
            self.muestra_mensaje(str(ex), 'red')
            return False

    def borrar_antiguos(self):
        """ Revisa la lista de archivos cargados y a partir de ella elimina los archivos correspondientes
            a versiones antiguas de un mismo libro """
        self.BarraBloque.setValue(0)
        self.muestra_mensaje('Comenzando búsqueda de versiones antiguas...', 'black')
        if self.radioMover.isChecked():
            self.muestra_mensaje('Las versiones antiguas se moverán a ' + self.dir_duplicados)
        self.num_borrados = 0
        actual = 0
        try:
            while actual < len(self.lista_archivos):
                QtWidgets.QApplication.processEvents()
                coincidencias = [actual]
                if (self.CoincSinId.isChecked()) or ('ID' in self.lista_archivos[actual]):
                    for i in range(actual + 1, len(self.lista_archivos) - 1):
                        if ('ext' in self.lista_archivos[actual] and 'ext' in self.lista_archivos[i]
                                and self.lista_archivos[actual]['ext'] == self.lista_archivos[i]['ext']):
                            if ('ID' in self.lista_archivos[i] and 'ID' in self.lista_archivos[actual]
                                    and self.lista_archivos[actual]['ID'] == self.lista_archivos[i]['ID']):
                                if self.RevisarTitulos.isChecked():
                                    if self.comparar_titulo_autor(self.lista_archivos[actual]['titulo_comp'], '',
                                                                  self.lista_archivos[i]['titulo_comp'], ''):
                                        coincidencias.append(i)
                                    else:
                                        elegido = QtWidgets.QMessageBox.question(None, "Coincidencia parcial",
                                                                                 "Los siguientes archivos tienen el mismo ID (%s), pero su título no coincide:\r\n" %
                                                                                 self.lista_archivos[actual]['ID']
                                                                                 + self.lista_archivos[actual][
                                                                                     'titulo_comp'] + "\r\n" +
                                                                                 self.lista_archivos[i][
                                                                                     'titulo_comp'] + "\r\n"
                                                                                 + "¿Se tratan en realidad del mismo libro?",
                                                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                                                 QtWidgets.QMessageBox.No)
                                        if elegido == QtWidgets.QMessageBox.Yes:
                                            coincidencias.append(i)
                                else:
                                    coincidencias.append(i)
                            elif self.CoincSinId.isChecked() and 'ID' not in self.lista_archivos[i] or 'ID' not in \
                                    self.lista_archivos[actual]:
                                if self.lista_archivos[actual]['titulo_comp'] == self.lista_archivos[i]['titulo_comp']:
                                    coincidencias.append(i)
                # Si hay 2 o más coincidencias, procedo a borrar
                if len(coincidencias) >= 2:
                    actual = self.comparar_y_borrar(coincidencias, actual)
                else:
                    actual += 1
                # Actualizo el número de pasos a las barras total y parcial
                self.BarraTotal.setMaximum(2 * len(self.lista_archivos))
                self.BarraBloque.setMaximum(len(self.lista_archivos))
                # Actualización de barras de progreso
                if actual <= self.BarraBloque.maximum():
                    self.BarraBloque.setValue(actual)
                prog_max = round(self.BarraTotal.maximum() / 2) + actual
                if prog_max <= self.BarraTotal.maximum():
                    self.BarraTotal.setValue(prog_max)
                # Comprobación de borrados
                if self.cancelar_borrado:
                    self.muestra_mensaje('Proceso de búsqueda cancelado', 'darkgreen')
                    return
        except Exception as ex:
            self.muestra_mensaje('Se ha producido el siguiente error durante la búsqueda de las versiones antiguas:',
                                 'red')
            self.muestra_mensaje(str(ex), 'red')
        if self.radioEliminar.isChecked():
            self.muestra_mensaje(
                'Búsqueda de versiones antiguas finalizada (' + str(self.num_borrados) + ' archivos eliminados)',
                'darkgreen')
        else:
            self.muestra_mensaje(
                'Búsqueda de versiones antiguas finalizada (' + str(self.num_borrados) + ' archivos movidos)',
                'darkgreen')

    def comparar_y_borrar(self, coincidencias, ind_actual):
        # Revisa las coincidencias encontradas en la lista de archivos y borra las versiones antiguas
        elem_borrar = []
        ind_max = coincidencias[0]
        # Localización de los que hay que borrar
        for indice in coincidencias[1:]:
            if ('version' in self.lista_archivos[indice]) and ('version' in self.lista_archivos[ind_max]):
                if (self.lista_archivos[indice]['version'] > self.lista_archivos[ind_max]['version']):
                    elem_borrar.append(ind_max)
                    ind_max = indice
                elif (self.lista_archivos[indice]['version'] < self.lista_archivos[ind_max]['version']):
                    elem_borrar.append(indice)
        # Borrado de elementos
        for i in range(len(elem_borrar) - 1, -1, -1):
            # Borrado en disco
            try:
                nom_archivo = os.path.join(self.dir_biblio, self.lista_archivos[elem_borrar[i]]['nombre'])
                if sys.platform == "win32":
                    os.chmod(nom_archivo, stat.S_IWRITE)
                if self.radioEliminar.isChecked():
                    # Borrado
                    os.remove(nom_archivo)
                    self.muestra_mensaje(
                        'Se ha eliminado correctamente el archivo ' + self.lista_archivos[elem_borrar[i]]['nombre'],
                        'blue')
                else:
                    # Traslado a otro directorio
                    destino = os.path.join(self.dir_duplicados, self.lista_archivos[elem_borrar[i]]['nombre'])
                    if os.path.isfile(destino):
                        os.remove(destino)
                    shutil.move(nom_archivo, destino)
                    self.muestra_mensaje(
                        'Se ha movido correctamente el archivo ' + self.lista_archivos[elem_borrar[i]]['nombre'],
                        'blue')
                self.num_borrados += 1
            except Exception as ex:
                self.muestra_mensaje('Se ha producido un error al intentar eliminar o mover el archivo ' +
                                     self.lista_archivos[elem_borrar[i]]['nombre'] + ':', 'red')
                self.muestra_mensaje(str(ex), 'red')
            del self.lista_archivos[elem_borrar[i]]
        # Aumento del índice actual global sólo si no se ha borrado el elemento actual
        if not (coincidencias[0] in elem_borrar):
            ind_actual += 1
        return ind_actual

    def filtrar_libros(self):
        # Carga la lista de libros en el directorio de la biblioteca, los compara con la lista de libros cargados
        # y elimina los que ya están descargados
        self.muestra_mensaje('Comenzando proceso de filtrado de enlaces...', 'black')
        exito = self.leer_directorio(True)
        if exito:
            self.filtrado = True
            mensaje = 'Realizando filtrado de enlaces comparándolos con los ya descargados...\n'
            mensaje += 'Puede llevar un rato, especialmente trabajando con el RSS o CSV total. Al principio va muy lento, pero luego acelera (prometido) ;)'
            self.muestra_mensaje(mensaje, 'black')
            exito = self.filtrar_lista()
            if exito:
                self.muestra_mensaje('Filtrado de enlaces realizado con éxito', 'darkgreen')
                return True
            else:
                self.muestra_mensaje('Se produjo un error durante el filtrado de enlaces', 'red')
                return False
        else:
            return False

    def filtrar_lista(self):
        try:
            i = 0
            # Doy el número de pasos a la barra parcial
            self.BarraBloque.setMaximum(len(self.lista_libros))
            self.BarraBloque.setValue(0)
            for i in range(len(self.lista_libros) - 1, -1, -1):
                for j in range(len(self.lista_archivos) - 1, -1, -1):
                    # Primero procesamos los que tienen ID
                    if ('epl_id' in self.lista_libros[i]) and ('ID' in self.lista_archivos[j]):
                        if self.lista_libros[i]['epl_id'] == self.lista_archivos[j]['ID']:
                            if ('version' in self.lista_libros[i] and 'version' in self.lista_archivos[j] and
                                    self.lista_libros[i]['version'] == self.lista_archivos[j]['version']):
                                if self.comparar_titulo_autor(self.lista_libros[i]['titulo'],
                                                              self.lista_libros[i]['autor'],
                                                              self.lista_archivos[j]['titulo'],
                                                              self.lista_archivos[j]['autor']):
                                    del self.lista_libros[i]
                                    del self.lista_archivos[j]
                                    break
                                else:
                                    msj = 'Se encontraron dos libros con el mismo ID, pero cuyos títulos y/o autores no coinciden exactamente'
                                    msj += ', por lo que el enlace no será filtrado:\n'
                                    msj += '[%s]\n' % self.lista_libros[i]['epl_id']
                                    msj += 'En RSS/CSV : ' + self.lista_libros[i]['autor'] + ' - ' + \
                                           self.lista_libros[i]['titulo'] + '\n'
                                    msj += 'Descargado : ' + self.lista_archivos[j]['autor'] + ' - ' + \
                                           self.lista_archivos[j]['titulo']
                                    self.muestra_mensaje(msj, 'brown')
                    # Ahora vamos con los que no tienen ID
                    if ('version' in self.lista_libros[i] and 'version' in self.lista_archivos[j] and
                            self.lista_libros[i]['version'] == self.lista_archivos[j]['version']):
                        if (self.comparar_titulo_autor(self.lista_libros[i]['titulo'], self.lista_libros[i]['autor'],
                                                       self.lista_archivos[j]['titulo'],
                                                       self.lista_archivos[j]['autor'])):
                            del self.lista_libros[i]
                            del self.lista_archivos[j]
                            break
                # Actualizo barra parcial
                self.BarraBloque.setMaximum(len(self.lista_libros))
                self.BarraBloque.setValue(len(self.lista_libros) - i)
                QtWidgets.QApplication.processEvents()
            return True
        except:
            return False

    def normalizar(self, cadena):
        # Normaliza una cadena de texto dejando sólo caracteres ASCII, paréntesis y guiones
        buscar = "ÃÀÁÄÂĀÅĂĄǍǞǠǺȦǢÈÉËÊĒĔĖĘĚÌÍÏÎĪĨĬĮİǏÒÓÖÔŌŎŐƠǑǪǬȪȬȮȰÙÚÜÛŪǕŨŬŮŰŲƯǓǕǗǙǛãàáäâāåăąǎǟǡǻȧǣèéëêēĕėęěìíïîīĩĭįǐòóöôōŏőơǒǫǭȫȭȯȱøùúüûūǖũŭůűųưǔǖǘǚǜÑñÇçłȲȳ-"
        sustit = "AAAAAAAAAAAAAAAEEEEEEEEEIIIIIIIIIIOOOOOOOOOOOOOOOUUUUUUUUUUUUUUUUUaaaaaaaaaaaaaaaeeeeeeeeeiiiiiiiiioooooooooooooooouuuuuuuuuuuuuuuuunncclYy "
        cadena = cadena.replace("´", "")
        for i in range(len(buscar)):
            cadena = cadena.replace(buscar[i], sustit[i])
        cadena = normalize('NFKD', cadena).encode('ascii', 'ignore').decode('utf-8', 'ignore')
        cadena = re.sub('[^A-Za-z0-9\s\&]+', '', cadena).strip()
        cadena = re.sub('\s+', ' ', cadena)
        return cadena

    def comparar_titulo_autor(self, titulo1, autor1, titulo2, autor2):
        # Comparar títulos y autores, para ver si coinciden ambos
        titulo1 = re.sub(r'\s0+(\d)', r' \1', titulo1.lower())
        titulo2 = re.sub(r'\s0+(\d)', r' \1', titulo2.lower())
        autor1 = autor1.lower()
        autor2 = autor2.lower()
        if (titulo1 == titulo2) and (autor1 == autor2):
            return True
        if not (titulo1 in titulo2) and not (titulo2 in titulo1):
            return False
        if not (autor1 in autor2) and not (autor2 in autor1):
            if ' & ' in autor1:
                coinc = True
                c_aut1 = autor1.split(' & ')
                for aut1 in c_aut1:
                    if not aut1 in autor2:
                        coinc = False
                if not coinc:
                    c_aut2 = autor2.split(' & ')
                    for aut2 in c_aut2:
                        if not aut2 in autor1:
                            return False
            else:
                return False
        return True


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)

    # Traduccion widget qt
    locale = QLocale.system().name()
    qtTranslator = QTranslator()
    if ((qtTranslator.load("qt_" + locale, QLibraryInfo.location(QLibraryInfo.TranslationsPath)))
            or (qtTranslator.load("qt_" + locale, directorio_ex()))):
        app.installTranslator(qtTranslator)

    window = Servidor()
    sys.exit(app.exec_())
