# Gestión de Pacientes, Citas y Facturación en Excel (VBA)

Este proyecto utiliza macros de VBA para gestionar la información de pacientes, citas médicas y facturación en un entorno de Excel. Está diseñado para automatizar procesos clave en una clínica o centro médico, facilitando la organización de datos y mejorando la eficiencia en tareas como la asignación de números de paciente, validación de DNI, búsqueda de citas y facturas, y registro de derivaciones clínicas.

## **Características principales**

### **1. Validación de DNI**
- **Función:** `ValidateDNI`
- Permite verificar si un DNI o NIE es válido.
- Compatible con documentos de extranjeros (letras iniciales `X`, `Y`, `Z`).
- Realiza el cálculo de la letra final del documento.

### **2. Gestión de pacientes**
- **Asignación automática de números de paciente:**
  - **Subrutina:** `AutomaticallyAssign`
  - Genera automáticamente un nuevo número de paciente basado en los registros existentes.
- **Determinación del último número de paciente:**
  - **Función:** `GetLastPatientNumber`
  - Busca el mayor número de paciente en varias hojas.

### **3. Derivaciones clínicas**
- **Función:** `ReferralClinic`
- Permite registrar información sobre la clínica de derivación de un paciente mediante un cuadro de entrada.

### **4. Gestión de facturación**
- **Búsqueda de facturas:**
  - **Subrutina:** `SearchBill`
  - Localiza facturas basadas en una fecha específica y las organiza según la especialidad.
- **Generación de número de factura:**
  - **Función:** `ObtenerNumeroFactura`
  - Crea un nuevo número de factura incrementando el último existente y agregando el año actual.
- **Guardar facturas:**
  - **Función:** `SaveBill`
  - Registra datos de una factura en una hoja de registro, con manejo de errores.

### **5. Gestión de citas**
- **Búsqueda de citas por día:**
  - **Subrutina:** `FinderDayAppointments`
  - Busca y organiza las citas programadas para una fecha específica.
- **Distribución de citas en horario:**
  - **Función:** `AppointmentFinder`
  - Asigna las citas a intervalos horarios específicos y aplica formato para resaltar datos clave (e.g., número de paciente, DNI, especialidad).

---

## **Estructura del proyecto**

El proyecto está dividido en las siguientes funciones y subrutinas:

| **Nombre**             | **Descripción**                                                                                       |
|------------------------|-------------------------------------------------------------------------------------------------------|
| `ValidateDNI`          | Valida un DNI o NIE según las reglas oficiales españolas.                                             |
| `AutomaticallyAssign`  | Asigna un número de paciente automáticamente si no existe uno registrado.                             |
| `GetLastPatientNumber` | Obtiene el último número de paciente registrado en las hojas de Excel.                                |
| `ReferralClinic`       | Registra la clínica de derivación de un paciente.                                                    |
| `SearchBill`           | Busca una factura en función de una fecha y organiza los datos en hojas específicas.                  |
| `ObtenerNumeroFactura` | Genera un nuevo número de factura único basado en el año actual.                                      |
| `SaveBill`             | Guarda los datos de una factura en una hoja de registro.                                             |
| `FinderDayAppointments`| Encuentra las citas de un día específico y organiza los datos en una hoja destino.                    |
| `AppointmentFinder`    | Procesa las citas encontradas, asignándolas a intervalos horarios específicos con formato destacado.  |

---

## **Requisitos del sistema**

1. Microsoft Excel (compatible con versiones que soporten VBA, e.g., Excel 2013 o superior).
2. Habilitación de macros en el libro de Excel.
3. Organización de las hojas en el archivo de Excel según las referencias del código:
   - `Hoja1`: Datos generales.
   - `Hoja2`, `Hoja3`, `Hoja4`: Registro de pacientes.
   - `Hoja5`: Facturas para especialidad de "Radiología".
   - `Hoja6`: Facturas para otras especialidades.
   - `Hoja7`: Panel de citas del día.
   - `Hoja8`: Registro de citas.
   - `Hoja9`: Registro de facturación.

---

## **Instrucciones de uso**

1. Descarga el archivo Excel con las macros habilitadas.
2. Asegúrate de que las hojas mencionadas estén configuradas correctamente en el libro.
3. Habilita las macros en Excel:
   - Ve a **Archivo > Opciones > Centro de confianza > Configuración de macros**.
   - Habilita todas las macros.
4. Ejecuta las macros según sea necesario:
   - Presiona `Alt + F8` en Excel para seleccionar y ejecutar las macros.
5. Interactúa con los cuadros de entrada y cuadros de mensaje para completar tareas como:
   - Validar DNI.
   - Asignar números de paciente.
   - Registrar derivaciones clínicas.
   - Buscar facturas y citas.

---

Si tienes preguntas o encuentras problemas, no dudes en abrir un *issue* en este repositorio. ¡Gracias por usar este proyecto! 🎉
