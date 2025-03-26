from fpdf import FPDF
import pandas as pd
from datetime import datetime, time, timedelta
import locale
import os

try:
    locale.setlocale(locale.LC_TIME, 'spanish')
except:
    pass

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'INFORME DE ASISTENCIA MENSUAL', 0, 1, 'C')
        self.ln(5)
        
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def parse_time(t):
    try:
        return datetime.strptime(str(t), "%H:%M:%S").time()
    except:
        return None

def generar_pdf(informe, nombre, rut, filename, total_atraso):
    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Encabezado
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, f'Nombre: {nombre}', 0, 1)
    pdf.cell(0, 10, f'RUT: {rut}', 0, 1)
    pdf.ln(10)
    
    # Tabla
    columnas = ['Fecha', 'Día', 'Entrada', 'Salida', 'Duración', 'Horas Extras', 'Estado']
    anchos = [25, 25, 20, 20, 25, 25, 30]
    
    # Cabecera
    pdf.set_font('Arial', 'B', 10)
    for ancho, columna in zip(anchos, columnas):
        pdf.cell(ancho, 10, columna, border=1, align='C')
    pdf.ln()
    
    # Datos
    pdf.set_font('Arial', '', 9)
    for registro in informe:
        for ancho, campo in zip(anchos, [registro[c] for c in columnas]):
            pdf.cell(ancho, 10, str(campo), border=1)
        pdf.ln()
    
    #resumen
    pdf.set_font('Arial', 'B', 9)
    # Fila de total
    pdf.cell(115, 10, "Total de minutos de atraso:", border=1, align='R')
    pdf.cell(25, 10, str(total_atraso), border=1)
    pdf.ln()

    # Firmas
    pdf.ln(20)
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(95, 10, "Firma Trabajador:", 0)
    pdf.cell(95, 10, "Firma Supervisor:", 0)
    pdf.ln(15)
    pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x()+80, pdf.get_y())
    pdf.line(pdf.get_x()+105, pdf.get_y(), pdf.get_x()+185, pdf.get_y())
    
    pdf.output(filename)

def process_file(file_path):
        try:
            # Leer el archivo seleccionado
            df = pd.read_excel(file_path, skipfooter=1)
            
            # Obtener datos del trabajador
            nombre_persona = df['Nombre'].iloc[0]
            rut_persona = df['Rut'].iloc[0]
            
            # Convertir y procesar fechas
            df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d-%m-%Y', errors='coerce')
            df = df[df['Fecha'].notna()].copy()

            informe = []

            total_atraso = 0

            for index, row in df.iterrows():
                fecha = row['Fecha']
                if pd.isna(fecha):
                    nombre_dia = "Fecha inválida"
                    fecha_formateada = "N/A"
                else:
                    nombre_dia = fecha.strftime('%A').capitalize()
                    fecha_formateada = fecha.strftime('%d-%m-%Y')
                
                entrada = parse_time(row['Entrada']) if pd.notna(row['Entrada']) else None
                salida = parse_time(row['Salida']) if pd.notna(row['Salida']) else None
                
                minutos_atraso = 0
                salida_base = time(16, 0) if nombre_dia == "Viernes" else time(17, 0)
                salida_esperada = salida_base
                estado = "S/D"
                horas_extras = "00:00"
                duracion_real = "No calculable"

                # Cálculo de atrasos
                if entrada:
                    # Calcular minutos de atraso
                    if entrada > time(8, 0):
                        delta = datetime.combine(datetime.min, entrada) - datetime.combine(datetime.min, time(8, 0))
                        minutos_atraso = delta.seconds // 60
                        
                        # Lógica de salida esperada
                        if entrada > time(8, 15):
                            estado = f"{minutos_atraso} min (E/T)"
                            total_atraso += minutos_atraso #acumular total
                            salida_esperada = salida_base
                        else:
                            salida_esperada = (datetime.combine(datetime.min, salida_base) + 
                                            timedelta(minutes=minutos_atraso)).time()
                            
                            if salida and salida < salida_esperada:
                                delta_salida = datetime.combine(datetime.min, salida_esperada) - datetime.combine(datetime.min, salida)
                                total_atraso += (delta_salida.seconds/60)
                                estado = f"{delta_salida.seconds // 60} min (N/R)"
                
                # Cálculo de horas extras (desde salida ESPERADA)
                if salida and salida > salida_esperada:  # Cambio clave aquí
                    extras = datetime.combine(datetime.min, salida) - datetime.combine(datetime.min, salida_esperada)
                    horas = extras.seconds // 3600
                    minutos = (extras.seconds // 60) % 60
                    horas_extras = f"{horas:02d}:{minutos:02d}"
                else:
                    horas_extras = "00:00"
                
                # Cálculo de duración real
                if entrada and salida:
                    tiempo_real = datetime.combine(datetime.min, salida) - datetime.combine(datetime.min, entrada)
                    horas_total = tiempo_real.seconds // 3600
                    minutos_total = (tiempo_real.seconds // 60) % 60
                    duracion_real = f"{horas_total:02d}:{minutos_total:02d}"

                informe.append({
                    'Fecha': fecha_formateada,
                    'Día': nombre_dia,
                    'Entrada': entrada.strftime("%H:%M:%S") if entrada else "Sin registro",
                    'Salida': salida.strftime("%H:%M:%S") if salida else "Sin registro",
                    'Duración': duracion_real,
                    'Horas Extras': horas_extras,
                    'Estado': estado
                })

            # Generar PDF
            output_dir = os.path.dirname(file_path)
            output_name = f"Informe_{nombre_persona.replace(' ', '_')}.pdf"
            output_path = os.path.join(output_dir, output_name)
            
            generar_pdf(informe, nombre_persona, rut_persona, output_path, total_atraso)
            
        except Exception as e:
            raise RuntimeError(f"Error procesando archivo: {str(e)}")