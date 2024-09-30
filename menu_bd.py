import uno
import traceback
import subprocess
import psutil
import time

# -- FUNCION PARA COMPROBAR SI EL SERVIDOR LIBREOFFICE ESTA EJECUTANDOSE
def esta_activo_servidor():
    for proc in psutil.process_iter(['name']):
        try:
            if 'soffice' in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

# -- FUNCION PARA INICIAR EL SERVIDOR LIBRE OFFICE
def iniciar_el_servidor():
    try:
        subprocess.Popen([
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            '--headless',
            '--accept=socket,host=localhost,port=2002;urp;',
            '--norestore'
        ])
        print("Servidor de LibreOffice iniciado.")
    except Exception as e:
        print(f"Error al iniciar LibreOffice: {e}")
        traceback.print_exc()

# -- FUNCION PARA CERRAR LIBRE OFFICE USANDO API UNO
def para_api_uno():
    print("Intentando cerrar el servidor usando UNO.....")
    try:
        # Establecer la conexión al servidor UNO de LibreOffice
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        )
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        
        # Obtener el ServiceManager
        smgr = ctx.ServiceManager
        
        # Obtener el Desktop
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        
        # Cerrar todos los documentos abiertos
        components = desktop.getComponents()
        enumerator = components.createEnumeration()
        while enumerator.hasMoreElements():
            component = enumerator.nextElement()
            # Cerrar el documento
            try:
                component.dispose()
            except Exception as e:
                print(f"Error al cerrar un componente: {e}")
        
        # Terminar LibreOffice
        desktop.terminate()
        print("Se ha enviado la orden de cierre a LibreOffice.")
        
    except Exception as e:
        print(f"Error al intentar cerrar LibreOffice vía UNO: {e}")
        traceback.print_exc()

    # Comprobar si LibreOffice sigue corriendo después del cierre
    time.sleep(5)
    if esta_activo_servidor():
        print("LibreOffice sigue en ejecución. Intentando cierre forzado con pkill...")
        try:
            subprocess.run(["pkill", "soffice"])
            print("Cierre forzado de LibreOffice completado.")
        except Exception as e:
            print(f"Error al forzar el cierre de LibreOffice: {e}")
            traceback.print_exc()
    else:
        print("LibreOffice cerrado correctamente.")

# -- FUNCION PARA CONECTAR AL SERVIDOR LIBREOFFICE
def conectar_servidor():
    try:
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        )
        time.sleep(10)
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

        # Obtener el ServiceManager
        smgr = ctx.ServiceManager

        # Conectarse a la base de datos a través de LibreOffice Base
        dbContext = smgr.createInstanceWithContext("com.sun.star.sdb.DatabaseContext", ctx)

        # Ruta al archivo .odb en formato URL
        ruta_bd = input("Introduzca la ruta a la base de datos. formato /....: ")
        odb_path = f'file://{ruta_bd}'

        # Conectar a la base de datos LibreOffice
        dataSource = dbContext.getByName(odb_path)
        connection = dataSource.getConnection("", "")
        
        return connection, dataSource, smgr
    except Exception as e:
        print(f"Error al conectar con la base de datos: {e}")
        traceback.print_exc()
        raise

# -- FUNCION PARA LISTAR TABLAS DE LA BASE DE DATOS USANDO METADATOS
def listar_tablas_metadata(connection):
    try:
        # Obtener los metadatos de la conexión
        metaData = connection.getMetaData()
        # Usar cadenas vacías en lugar de None
        resultSet = metaData.getTables("", "", "%", ("TABLE",))
        
        print("\n--- Lista de Tablas ---")
        tablas = []
        while resultSet.next():
            table_name = resultSet.getString("TABLE_NAME")
            print(f"- {table_name}")
            tablas.append(table_name)
        resultSet.close()
        return tablas
    except Exception as e:
        print(f"Error al listar las tablas usando metadatos: {e}")
        traceback.print_exc()
        return []

# -- FUNCION PARA LISTAR TABLAS DE LA BASE DE DATOS USANDO dataSource
def listar_tablas_datasource(dataSource):
    try:
        # Acceder directamente a las tablas desde dataSource
        tablas = []
        print("\n--- Lista de Tablas ---")
        for table in dataSource.DatabaseDocument.DataSource.Tables:
            table_name = table.Name
            print(f"- {table_name}")
            tablas.append(table_name)
        return tablas
    except Exception as e:
        print(f"Error al listar las tablas usando dataSource: {e}")
        traceback.print_exc()
        return []

# -- FUNCION PARA LISTAR TABLAS (INTELIGENTE)
def listar_tablas(connection, dataSource):
    # Intentar primero con dataSource
    tablas = listar_tablas_datasource(dataSource)
    if not tablas:
        # Si falla, intentar con metadata
        tablas = listar_tablas_metadata(connection)
    return tablas

# -- FUNCION PARA MOSTRAR COLUMNAS DE UNA TABLA ESPECIFICA
def mostrar_columnas(connection, tabla):
    try:
        statement = connection.createStatement()
        query = f'SELECT * FROM "{tabla}" LIMIT 1'
        resultSet = statement.executeQuery(query)
        
        metaData = resultSet.MetaData
        columnCount = metaData.ColumnCount

        print(f"\n--- Columnas de la tabla '{tabla}' ---")
        for i in range(1, columnCount + 1):
            print(f"Columna {i}: {metaData.getColumnName(i)}")
        
        resultSet.close()
        statement.close()
    except Exception as e:
        print(f"Error al mostrar las columnas de la tabla {tabla}: {e}")
        traceback.print_exc()

# -- FUNCION PARA MOSTRAR CONTENIDO DE UNA TABLA
def mostrar_contenido(connection, tabla):
    try:
        statement = connection.createStatement()
        query = f'SELECT * FROM "{tabla}"'
        resultSet = statement.executeQuery(query)
        
        metaData = resultSet.MetaData
        columnCount = metaData.ColumnCount

        # Obtener nombres de columnas
        columnas = [metaData.getColumnName(i) for i in range(1, columnCount + 1)]
        print(f"\n--- Contenido de la tabla '{tabla}' ---")
        print("\t".join(columnas))
        
        # Obtener y mostrar filas
        while resultSet.next():
            fila = [resultSet.getString(i) if resultSet.getString(i) is not None else "" for i in range(1, columnCount + 1)]
            print("\t".join(fila))
        
        resultSet.close()
        statement.close()
    except Exception as e:
        print(f"Error al mostrar el contenido de la tabla {tabla}: {e}")
        traceback.print_exc()

# -- FUNCION PARA MOSTRAR EL MENU
def mostrar_menu():
    print("\n=== Menú Principal ===")
    print("1. Listar Tablas de la Base de Datos")
    print("2. Mostrar Columnas de una Tabla")
    print("3. Mostrar Contenido de una Tabla")
    print("4. Salir")

# -- FUNCION PRINCIPAL PARA EJECUTAR EL PROGRAMA
def main():
    # Verificar si el servidor de LibreOffice está en ejecución
    server_started = False
    connection = None  # Inicializar la variable
    
    if not esta_activo_servidor():
        print("Iniciando el servidor de LibreOffice...")
        iniciar_el_servidor()
        time.sleep(5)  # Esperar a que el servidor inicie
        server_started = True
        # Verificar nuevamente si el servidor está activo
        if not esta_activo_servidor():
            print("No se pudo iniciar el servidor de LibreOffice. Saliendo del programa.")
            return


    try:
        # Conectar a la base de datos y obtener el dataSource y smgr
        connection, dataSource, smgr = conectar_servidor()
        print("Conexión a la base de datos establecida.")

        while True:
            mostrar_menu()
            opcion = input("Selecciona una opción (1-4): ").strip()

            if opcion == '1':
                # Listar tablas
                tablas = listar_tablas(connection, dataSource)
                if not tablas:
                    print("No se encontraron tablas en la base de datos.")
            elif opcion == '2':
                # Mostrar columnas de una tabla
                tablas = listar_tablas(connection, dataSource)
                if not tablas:
                    print("No hay tablas disponibles para mostrar columnas.")
                    continue
                tabla = input("Ingrese el nombre de la tabla para mostrar sus columnas: ").strip()
                if tabla not in tablas:
                    print(f"La tabla '{tabla}' no existe. Por favor, selecciona una tabla válida.")
                else:
                    mostrar_columnas(connection, tabla)
            elif opcion == '3':
                # Mostrar contenido de una tabla
                tablas = listar_tablas(connection, dataSource)
                if not tablas:
                    print("No hay tablas disponibles para mostrar contenido.")
                    continue
                tabla = input("Ingrese el nombre de la tabla para mostrar su contenido: ").strip()
                if tabla not in tablas:
                    print(f"La tabla '{tabla}' no existe. Por favor, selecciona una tabla válida.")
                else:
                    mostrar_contenido(connection, tabla)
            elif opcion == '4':
                # Salir del programa
                print("Saliendo del programa...")
                break
            else:
                print("Opción inválida. Por favor, selecciona una opción entre 1 y 4.")

    except Exception as e:
        print(f"Error durante la operación: {e}")
        traceback.print_exc()

    finally:
        # Cerrar la conexión a la base de datos
        try:
            if connection is not None:
                connection.close()
                print("Conexión a la base de datos cerrada.")
        except Exception as e:
            print(f"Error al cerrar la conexión: {e}")
            traceback.print_exc()
        
        # Cerrar LibreOffice de manera ordenada
        if server_started:
            para_api_uno()


if __name__ == "__main__":
    main()
