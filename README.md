- Chatbot de Gestión de Vacaciones

- Descripción

Trabajo Práctico Integrador de la materia Organización Empresarial.

El proyecto consiste en la automatización del proceso de solicitud de vacaciones mediante un chatbot desarrollado en Python. El sistema consulta una base de datos simulada en Excel, valida reglas de negocio definidas en el modelo BPMN y registra cada solicitud realizada por los empleados.

- Funcionalidades

* Solicitud de vacaciones.
* Consulta de saldo disponible.
* Validación de formato de fechas.
* Validación de fechas retroactivas.
* Control de saldo de días disponibles.
* Registro de solicitudes aprobadas, rechazadas y canceladas.
* Actualización automática del saldo de vacaciones.

- Tecnologías Utilizadas

* Python 3
* OpenPyXL
* Excel (.xlsx)
* Git
* GitHub

- Estructura del Proyecto

text
TPI.py
base_vacaciones.xlsx
input.txt
README.md


- Instalación

Instalar la dependencia necesaria:

bash
pip install openpyxl


- Ejecución

Ejecutar el programa con:

bash
python TPI.py


- Flujo de Funcionamiento

1. El usuario ingresa su ID.
2. Ingresa la fecha de inicio y fin de vacaciones.
3. El sistema valida el formato de las fechas.
4. Se verifica que las fechas sean válidas.
5. Se consulta el saldo disponible del empleado.
6. Se aprueba o rechaza la solicitud según las reglas de negocio.
7. La solicitud queda registrada en la base de datos simulada.

- Integrantes

* Integrante A: Análisis del proceso, BPMN y documentación.
* Integrante B: Desarrollo técnico, integración de datos y pruebas.

- Materia

Tecnicatura Universitaria en Programación a Distancia (TUPaD)

Organización Empresarial
