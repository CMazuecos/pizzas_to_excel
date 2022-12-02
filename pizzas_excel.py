import pandas as pd
import datetime
import re
import numpy as np
import xlsxwriter

def informe_calidad_datos (fichero, name):
    #Print name of the file, nans, nulls and data types
    print('Nombre del fichero:', name)
    print (fichero.isnull().sum()) 
    print (fichero.dtypes)
    print (fichero.isna().sum())
    datatype_dictionary = {}
    #escribir en un datatype_dictionary, como key el nombre de la columna y como value el tipo de dato
    for i in range(len(fichero.columns)):
        datatype_dictionary[fichero.columns[i]] = fichero.dtypes[i]
    #Number of nulls values
    print (fichero.isnull().sum())
    return datatype_dictionary
def limpiar_fichero_order_details(fichero):
    print(fichero)
    print('Limpiando fichero order_details')
    #we change nan values to 1
    fichero['quantity'].fillna(1, inplace=True)
    fichero['quantity'] = fichero['quantity'].apply(lambda quantity: change_quantity(quantity))
    #we change the nan values of the column pizza_id to the value above
    fichero['pizza_id'].fillna(method='ffill', inplace=True)
    fichero['pizza_id'] = fichero['pizza_id'].apply(lambda pizza_id: change_pizza_id(pizza_id))
    print('Fichero order_details limpio')
    return fichero  

def change_quantity(quantity):
    #change in the column quantity the values that are not numbers to their numbers and if they are less than 1 we change them to 1
    try:
        quantity = int(quantity)
        if quantity < 1:
            quantity = 1
        return quantity
    except:
        quantity = re.sub(r'one', '1', quantity, flags=re.IGNORECASE)
        quantity = re.sub(r'two', '2', quantity, flags=re.IGNORECASE)
        return quantity

def change_pizza_id(pizza_id):
    #convert some characters to others 
    pizza_id = re.sub('@', 'a', pizza_id)
    pizza_id = re.sub('0', 'o', pizza_id)
    pizza_id = re.sub('1', 'i', pizza_id)
    pizza_id = re.sub('3', 'e', pizza_id)
    pizza_id = re.sub('-', '_', pizza_id)
    pizza_id = re.sub(' ', '_', pizza_id)
    return pizza_id


def limpieza_datos_orders(fichero):
    fichero = fichero.sort_values(by=['order_id'])
    #fichero fillna with values above
    print(fichero)
    print('Limpiando fichero orders')
    fichero.fillna(method='ffill', inplace=True) # we fill the nan values with the value above
    #change column date to datetime
    fichero['date'] = fichero['date'].apply(lambda x: change_date(x))    
    print('Fichero orders limpio') 
    return fichero

def change_date(date):
    try:
        #change the date to datetime
        date = pd.to_datetime(date)
        return date
    except:
        #change the date (written in seconds) to the correct format
        date = datetime.datetime.fromtimestamp(int(float(date)))
        return date

def create_dictionary(pizza_types):
    #create a dictionary with the pizza type id as key and the ingredients as value 
    dictionary_pizza_type = {}
    for i in range(len(pizza_types)):
        dictionary_pizza_type [pizza_types ['pizza_type_id'][i]] = pizza_types ['ingredients'] [i]
    return dictionary_pizza_type

def cargar_datos (order_details, pizzas, pizza_types, orders):
    dictionary_pizza_type = create_dictionary(pizza_types)
    
    semanas, dias_semana = organizar_por_semanas(orders)
    pedidos, tamanos_cantidad, dinero = organizar_por_pedidos(semanas, order_details, dictionary_pizza_type, pizzas)
    ingredients_dictionary = {}
    for i in range(len(pedidos)):

        ingredients_week = transform_pizza_into_ingredients(pedidos[i], dias_semana[i], pizza_types, {})
        ingredients_dictionary [i+1] = ingredients_week
        print('Cargado los ingredientes de la semana', i+1)
    return ingredients_dictionary, pedidos, tamanos_cantidad, dinero

def organizar_por_semanas(orders):
    diccionario_weekdays = {}
    diccionario_pedidos = {}
    for i in range (53):
        diccionario_weekdays [i] = [0, 0, 0, 0, 0, 0, 0]
        diccionario_pedidos [i] = [] 

    for order in orders['order_id']: #add the order to the dictionary and the day of the week to the dictionary
        try:
            fecha = orders['date'][order]
            numero_semana = fecha.isocalendar().week
            numero_dia = fecha.isocalendar().weekday
            diccionario_weekdays [numero_semana-1][numero_dia-1] += 1
            diccionario_pedidos [numero_semana-1].append(orders['order_id'][order])
        except:
            pass
    for i in range(len(diccionario_weekdays)): #change the list to the number of days in the week
        dias_semana = 0
        for j in range(len(diccionario_weekdays[i])):
            if diccionario_weekdays[i][j] != 0:
                dias_semana += 1
        diccionario_weekdays[i] = dias_semana
    return diccionario_pedidos, diccionario_weekdays

def organizar_por_pedidos(semanas, order_details, dictionary_pizza_type, pizzas):
    tamanos = {'S': 1, 'M': 2, 'L': 3, 'XL': 4, 'XXL': 5}
    cantidad_tamanos = {'S': 0, 'M': 0, 'L': 0, 'XL': 0, 'XXL': 0}
    pedidos_semana = []
    dinero = []
    for i in range(len(semanas)):
        dinero_semana = 0
        pedidos_semana.append({})
        for key, value in dictionary_pizza_type.items():
            pedidos_semana[i][key] = 0
        

        for j in range(len(semanas[i])):
            #organize the weeks by orders
            order_id_buscado = semanas[i][j]
            lista_pizzas = order_details.loc[order_details['order_id'] == order_id_buscado]
            for pizza in lista_pizzas['pizza_id']:
                pizza_searched = pizzas.loc[pizzas['pizza_id'] == pizza]
                dinero_semana += pizza_searched['price'].values[0] #add the price of the pizza to the total money of the week
                quantity = lista_pizzas.loc[lista_pizzas['pizza_id'] == pizza]['quantity'].values[0]
                pizza_type = pizza_searched['pizza_type_id'].values[0]
                pizza_size = pizza_searched['size'].values[0]
                cantidad_tamanos[pizza_size] += 1 #add the size of the pizza to the dictionary
                pedidos_semana[i][pizza_type] += int(quantity) * int(tamanos[pizza_size])
        dinero_semana = round(dinero_semana, 2)
        dinero.append(dinero_semana)
        print('Cargado el pedido de la semana', i+1)
    return pedidos_semana, cantidad_tamanos, dinero

def transform_pizza_into_ingredients(pizzas_semana, dias_semana, pizza_types, ingredients_dictionary):    
    #get all the possible ingredients and add them to the dictionary
    for i in range(len(pizza_types)):
        ingredients = pizza_types['ingredients'][i]
        ingredients = ingredients.split(', ')
        for ingredient in ingredients:
            ingredients_dictionary[ingredient] = 0
    #add the ingredients of each pizza to the dictionary
    for key, value in pizzas_semana.items():
        ingredients = pizza_types.loc[pizza_types['pizza_type_id'] == key]['ingredients'].values[0]
        ingredients = ingredients.split(', ')
        for ingredient in ingredients:
            ingredients_dictionary[ingredient] += value
    for key, value in ingredients_dictionary.items():
        ingredients_dictionary[key] = int(np.ceil(value/dias_semana*7))
    return ingredients_dictionary

def extract_data():
    datatype_dictionary = {'order_details': {}, 'order_details': {}, 'pizza_types': {}, 'orders': {}}
    order_details = pd.read_csv('order_details.csv',sep=';')
    datatype_od = informe_calidad_datos(order_details, 'order_details.csv')
    datatype_dictionary['order_details'] = datatype_od
    order_details = limpiar_fichero_order_details(order_details)
    pizzas = pd.read_csv('pizzas.csv',sep = ',')
    datatype_p = informe_calidad_datos(pizzas, 'pizzas.csv')
    datatype_dictionary['pizzas'] = datatype_p
    pizza_types = pd.read_csv('pizza_types.csv', sep = ',', encoding='latin-1')
    datatype_pt = informe_calidad_datos(pizza_types, 'pizza_types.csv')
    datatype_dictionary['pizza_types'] = datatype_pt
    orders = pd.read_csv('orders.csv', sep = ';')
    datatype_o = informe_calidad_datos(orders, 'orders.csv')
    datatype_dictionary['orders'] = datatype_o
    orders = limpieza_datos_orders(orders)
    return order_details, pizzas, pizza_types, orders, datatype_dictionary

def calcular_pedidos_totales(pedidos):
    pedidos_totales = {}
    #get all the possible ingredients and add them to the dictionary
    for semana in pedidos:
        for key, value in semana.items():
            if key in pedidos_totales:
                pedidos_totales[key] += value
            else:
                pedidos_totales[key] = value
    return pedidos_totales

def load_data(ingredients, pedidos, tamanos_cantidad, dinero):
    #get all the orders and sort them by the number of orders
    pedidos_totales = calcular_pedidos_totales(pedidos)
    pedidos_sorted, pedido_value = [], []
    for key, value in pedidos_totales.items():
        pedidos_sorted.append(key)
        pedido_value.append(value)
    pedido_value.sort()
    pedidos_sorted.sort(key = lambda x: pedidos_totales[x])
    pedidos_sorted.reverse()
    pedido_value.reverse()
    #get the size of the pizzas quantities
    tamanos, cantidad = [], []
    for key, value in tamanos_cantidad.items():
        tamanos.append(key)
        cantidad.append(value)
    #get the ingredients and sort them by the quantity
    ingredientes_totales = {}
    for i in range(len(ingredients)):
        for key, value in ingredients[i+1].items():
            if key in ingredientes_totales:
                ingredientes_totales[key] += value
            else:
                ingredientes_totales[key] = value
    
    ingredientes_sorted, cantidad_sorted = [], []
    for key, value in ingredientes_totales.items():
        ingredientes_sorted.append(key)
        cantidad_sorted.append(value)
    cantidad_sorted.sort()
    ingredientes_sorted.sort(key = lambda x: ingredientes_totales[x])
    ingredientes_sorted.reverse()
    cantidad_sorted.reverse()
    #create a excel
    workbook = xlsxwriter.Workbook('Informe.xlsx')

    #FIRST SHEET

    worksheet = workbook.add_worksheet('Pedidos')
    #Make the second and third column wider
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(2, 2, 20)
    cell_format = workbook.add_format()
    cell_format.set_border()
    cell_format.set_bg_color('#ADD8E6') 
    cell_format.set_align('left')
    #make a table
    worksheet.write('B2', 'Pedidos', cell_format)
    worksheet.write('C2', 'Cantidad', cell_format)
    for i in range(len(pedidos_sorted)):
        worksheet.write(i+2, 1, pedidos_sorted[i], cell_format)
        worksheet.write(i+2, 2, pedido_value[i], cell_format)
    worksheet.write('B'+str(len(pedidos_sorted)+4), 'Tamaños', cell_format)
    worksheet.write('C'+str(len(pedidos_sorted)+4), 'Cantidad', cell_format)
    for i in range(len(tamanos)):
        worksheet.write(i+len(pedidos_sorted)+4, 1, tamanos[i], cell_format)
        worksheet.write(i+len(pedidos_sorted)+4, 2, cantidad[i], cell_format)
    worksheet.insert_image('E2', 'logo.jpg', {'x_scale': 2, 'y_scale': 2})
    #add a chart
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({
        'name': 'Pedidos',
        'categories': '=Pedidos!$B$3:$B$'+str(len(pedidos_sorted)+2),
        'values': '=Pedidos!$C$3:$C$'+str(len(pedido_value)+2),
    })
    
    chart1.set_title({'name': 'Pizzas pedidas'})
    chart1.set_x_axis({'name': 'Pedidos'})
    chart1.set_y_axis({'name': 'Cantidad'})
    chart1.set_style(11)
    worksheet.insert_chart('E15', chart1, {'x_offset': 50, 'y_offset': 50})
    #graficar los 10 pedidos mas solicitados
    top_10 = workbook.add_chart({'type': 'column'})
    top_10.add_series({
        'name': 'Pedidos',
        'categories': '=Pedidos!$B$3:$B$12',
        'values': '=Pedidos!$C$3:$C$12',
    })
    top_10.set_title({'name': 'Pizzas más pedidas'})
    top_10.set_x_axis({'name': 'Pedidos'})
    top_10.set_y_axis({'name': 'Cantidad'})
    top_10.set_style(11)
    worksheet.insert_chart('E30', top_10, {'x_offset': 50, 'y_offset': 50})
    #graficar los 10 pedidos menos solicitados
    worst_10 = workbook.add_chart({'type': 'column'})
    worst_10.add_series({
        'name': 'Pedidos',
        'categories': '=Pedidos!$B$'+str(len(pedidos_sorted)-8)+':$B$'+str(len(pedidos_sorted)+2),
        'values': '=Pedidos!$C$'+str(len(pedidos_sorted)-8)+':$C$'+str(len(pedidos_sorted)+2),
    })
    worst_10.set_title({'name': 'Pizzas menos pedidas'})
    worst_10.set_x_axis({'name': 'Pedidos'})
    worst_10.set_y_axis({'name': 'Cantidad'})
    worst_10.set_style(11)
    worksheet.insert_chart('E45', worst_10, {'x_offset': 50, 'y_offset': 50})

        
    
    
    pie_pizza = workbook.add_chart({'type': 'pie'})
    pie_pizza.add_series({
        'name': 'Tamaño de pizza',
        'categories': '=Pedidos!$B$'+str(len(pedidos_sorted)+4)+':$B$'+str(len(pedidos_sorted)+len(tamanos)+4),
        'values': '=Pedidos!$C$'+str(len(pedidos_sorted)+4)+':$C$'+str(len(pedidos_sorted)+len(tamanos)+4),
    })
    pie_pizza.set_title({'name': 'Tamaño de pizza'})
    pie_pizza.set_style(10)
    pie_pizza.set_legend({'position': 'bottom'})
    worksheet.insert_chart('E60', pie_pizza, {'x_offset': 50, 'y_offset': 50})
    
    #SECOND SHEET
    
    worksheet = workbook.add_worksheet('Ingredientes')
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 2, 30)
    cell_format = workbook.add_format()
    cell_format.set_border()
    cell_format.set_bg_color('#EF81F6') 
    cell_format.set_align('left')
    #hacer una tabla con los ingredientes y la cantidad de veces que se usaron
    worksheet.write('B2', 'Ingredientes', cell_format)
    worksheet.write('C2', 'Cantidad', cell_format)
    for i in range(len(ingredientes_sorted)):
        worksheet.write(i+2, 1, ingredientes_sorted[i], cell_format)
        worksheet.write(i+2, 2, cantidad_sorted[i], cell_format)
    # Add a chart
    ingredients_graph = workbook.add_chart({'type': 'column'})
    ingredients_graph.add_series({
        'name': 'Ingredientes',
        'categories': '=Ingredientes!$B$3:$B$'+str(len(ingredientes_sorted)+2),
        'values': '=Ingredientes!$C$3:$C$'+str(len(cantidad_sorted)+2),
    })
    ingredients_graph.set_title({'name': 'Ingredientes pedidos'})
    ingredients_graph.set_x_axis({'name': 'Ingredientes'})
    ingredients_graph.set_y_axis({'name': 'Cantidad'})
    ingredients_graph.set_style(11)
    worksheet.insert_chart('E15', ingredients_graph, {'x_offset': 50, 'y_offset': 50})
    worksheet.insert_image('E2', 'logo.jpg', {'x_scale': 2, 'y_scale': 2})
    #top 10 ingredients
    top_10_ingredients = workbook.add_chart({'type': 'column'})
    top_10_ingredients.add_series({
        'name': 'Ingredientes',
        'categories': '=Ingredientes!$B$3:$B$12',
        'values': '=Ingredientes!$C$3:$C$12',
    })
    top_10_ingredients.set_title({'name': 'Ingredientes más necesitado'})
    top_10_ingredients.set_x_axis({'name': 'Ingredientes'})
    top_10_ingredients.set_y_axis({'name': 'Cantidad'})
    top_10_ingredients.set_style(11)
    worksheet.insert_chart('E30', top_10_ingredients, {'x_offset': 50, 'y_offset': 50})
    #worst 10 ingredients
    worst_10_ingredients = workbook.add_chart({'type': 'column'})
    worst_10_ingredients.add_series({
        'name': 'Ingredientes',
        'categories': '=Ingredientes!$B$'+str(len(ingredientes_sorted)-8)+':$B$'+str(len(ingredientes_sorted)+2),
        'values': '=Ingredientes!$C$'+str(len(ingredientes_sorted)-8)+':$C$'+str(len(ingredientes_sorted)+2),
    })
    worst_10_ingredients.set_title({'name': 'Ingredientes menos necesitado'})
    worst_10_ingredients.set_x_axis({'name': 'Ingredientes'})
    worst_10_ingredients.set_y_axis({'name': 'Cantidad'})
    worst_10_ingredients.set_style(11)
    worksheet.insert_chart('E45', worst_10_ingredients, {'x_offset': 50, 'y_offset': 50})

    #THIRD SHEET
    worksheet = workbook.add_worksheet('Reporte_Ejecutivo')
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 2, 30)
    cell_format = workbook.add_format()
    cell_format.set_border()
    #añadir ingresos
    cell_format.set_bg_color('#6FFF6B')
    cell_format.set_align('left')
    worksheet.write('B2', 'Semana', cell_format)
    worksheet.write('C2', 'Ingresos', cell_format)
    for i in range(len(dinero)):
        worksheet.write(i+2, 1, i+1, cell_format)
        worksheet.write(i+2, 2, int(dinero[i]), cell_format)
    worksheet.insert_image('E2', 'logo.jpg', {'x_scale': 2, 'y_scale': 2})
    #graficar ingresos
    ingresos_graph = workbook.add_chart({'type': 'column'})
    ingresos_graph.add_series({
        'name': 'Ingresos',
        'categories': '=Reporte_Ejecutivo!$B$3:$B$'+str(len(dinero)+2),
        'values': '=Reporte_Ejecutivo!$C$3:$C$'+str(len(dinero)+2),
    })
    ingresos_graph.set_title({'name': 'Ingresos por semana'})
    ingresos_graph.set_x_axis({'name': 'Semana'})
    ingresos_graph.set_y_axis({'name': 'Ingresos'})
    ingresos_graph.set_style(11)
    worksheet.insert_chart('E15', ingresos_graph, {'x_offset': 50, 'y_offset': 50})


    workbook.close()






    





    





        
    
if __name__ == '__main__':
    order_details, pizzas, pizza_types, orders, datatype_dictionary = extract_data()
    ingredients, pedidos, tamanos_cantidad, dinero = cargar_datos(order_details, pizzas, pizza_types, orders)
    load_data(ingredients, pedidos, tamanos_cantidad, dinero)