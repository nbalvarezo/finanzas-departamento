import openpyxl
from datetime import datetime

# Función para ingresar datos financieros
def input_financial_data():
    # Ingresar salarios
    salary_person1 = float(input("Ingrese el salario de la persona 1: "))
    salary_person2 = float(input("Ingrese el salario de la persona 2: "))
    
    # Ingresar gastos
    expenses_person1 = float(input("Ingrese los gastos de la persona 1: "))
    expenses_person2 = float(input("Ingrese los gastos de la persona 2: "))
    
    # Ingresar alquiler
    rent = float(input("Ingrese el alquiler del departamento: "))
    
    return salary_person1, salary_person2, expenses_person1, expenses_person2, rent

# Función para crear y exportar archivo Excel
def export_to_excel(salary_person1, salary_person2, expenses_person1, expenses_person2, rent):
    # Crear un libro de trabajo
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    # Ingresar datos en las celdas
    sheet['A1'] = 'Salario Persona 1'
    sheet['B1'] = 'Salario Persona 2'
    sheet['C1'] = 'Gastos Persona 1'
    sheet['D1'] = 'Gastos Persona 2'
    sheet['E1'] = 'Alquiler'
    sheet['F1'] = 'Total Salarios'
    sheet['G1'] = 'Total Gastos'
    sheet['H1'] = 'Balance'
    
    sheet.append([salary_person1, salary_person2, expenses_person1, expenses_person2, rent])
    sheet.append(['=SUM(A2:B2)', '=SUM(C2:D2)+E2', '=G2-H2'])
    
    # Obtener la fecha actual
    current_date = datetime.now().strftime("%Y-%m-%d")
    filename = f'finanzas_departamento_{current_date}.xlsx'
    
    # Guardar el archivo
    workbook.save(filename)
    
    print(f"El archivo '{filename}' ha sido creado y exportado con éxito.")

# Función principal
def main():
    salary_person1, salary_person2, expenses_person1, expenses_person2, rent = input_financial_data()
    export_to_excel(salary_person1, salary_person2, expenses_person1, expenses_person2, rent)

# Ejecutar la función principal
if __name__ == "__main__":
    main()
