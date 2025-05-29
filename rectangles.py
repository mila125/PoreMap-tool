import turtle
import numpy as np

#numeros_derecha_struct = np.array([[]] * 5, dtype=object)
lados_rectangulos = np.array([])
numeros_derecha = np.array([])
numeros = np.array([])
coordinates = np.zeros(2)
lados_cuadrado = np.array([])
numeros_cuadrado = np.array([])
numeros_cuadrado_derecha = np.array([])
numeros_cuadrado= np.array([])
numeros_derecha_cuadrados= np.array([])
lados_ramas = np.array([])
secuencia = np.array([])
diagonal_flag=False
perimetro=0
rigth_j_max=0
numeros_ramas_derecha=  np.array([])
MAX_DEPTH=0
fibonacci_dual = []  # Array global para almacenar la secuencia combinada
color_palet = np.array([])
k = 0
okidoki="nothing"
occupied_positions = {}  # Un diccionario global para saber qué lugares están ocupados
def fibonacci_hasta_maximo(max_valor):
    """
    Genera una secuencia de Fibonacci hasta que el valor máximo sea alcanzado o superado.
    
    :param max_valor: El valor máximo que no debe superar la secuencia.
    :return: Lista con la secuencia de Fibonacci hasta el límite.
    """
    if max_valor <= 0:
        return [0]

    seq = [0, 1]
    while True:
        nuevo_valor = seq[-1] + seq[-2]
        if nuevo_valor > max_valor:  # Detenemos cuando el siguiente valor supera el límite
            break
        seq.append(nuevo_valor)

    return seq


def fibonacci_dual(max_x, max_y):
    """
    Genera dos secuencias de Fibonacci dentro de un mismo array, limitadas por max_x y max_y.
    - Elementos en índices pares corresponden a la distancia en X.
    - Elementos en índices impares corresponden a la distancia en Y.

    :param max_x: Valor máximo que la secuencia de X no debe superar.
    :param max_y: Valor máximo que la secuencia de Y no debe superar.
    :return: Lista con ambas secuencias intercaladas.
    """
    # Obtener las secuencias limitadas por el valor máximo
    seq_x = fibonacci_hasta_maximo(max_x)
    seq_y = fibonacci_hasta_maximo(max_y)

    # Intercalar ambas secuencias en un solo array
    resultado = []
    i, j = 0, 0

    while i < len(seq_x) or j < len(seq_y):
        if i < len(seq_x):
            resultado.append(seq_x[i])  # Elemento para X (posición par)
            i += 1
        if j < len(seq_y):
            resultado.append(seq_y[j])  # Elemento para Y (posición impar)
            j += 1

    return resultado
def draw_same_squares_rectangle(cantidad, lado_cuadrado, t, h, main_width, main_height, pose):
    global numeros_ramas_derecha
    global perimetro
    global numeros_derecha
    global coordinates
    global lados_rectangulos
    global lados_ramas
    global okidoki
    global occupied_positions  # Asegúrate de tener esto declarado globalmente e inicializado como {}

    if not (0 < cantidad < 10000 and perimetro < 10000):
        return

    espacio_del_cuadrado = int(perimetro / cantidad)
    i = 0  # vertical (y)
    j = 0  # horizontal (x)

    coordinates[0] =  - (lados_rectangulos[h+2] / 2)

    if okidoki == "okidoki":
        coordinates[1] = numeros_derecha[1] - lados_rectangulos[h + 1]
        okidoki = "nothing"

    def is_free(x, y):
        key = (int(x) // 10, int(y) // 10)
        return not occupied_positions.get(key, False)

    while i < lados_rectangulos[h + 1]:
        j = 0
        x = coordinates[0]
        y = coordinates[1]

        while j < lados_rectangulos[h]:
            max_attempts = 50
            attempts = 0
            success = False

            tx = x  # copia temporal de x para intentar encontrar lugar libre

            while attempts < max_attempts:
                if is_free(tx, y):
                    success = True
                    break
                tx -= 10
                attempts += 1

            if not success:
                print("No se encontró espacio libre para este cuadrado.")
                j += lado_cuadrado
                continue

            coordinates[0] = tx
            coordinates[1] = y

            draw_rectangle(t, lado_cuadrado, lado_cuadrado)

            # Marcar la celda como ocupada
            key = (int(tx) // 10, int(y) // 10)
            occupied_positions[key] = True

            coordinates[0] += lado_cuadrado
            j += lado_cuadrado

        i += lado_cuadrado

        coordinates[0] =  - (lados_rectangulos[h+2] / 2)
    #coordinates[1] =+(lados_rectangulos[h+3])
    numeros_derecha = np.append(numeros_derecha, coordinates)
def draw_diagonal_branch(t, width, height, angle):
    t.setheading(angle)
    t.begin_fill()
    for _ in range(2):
        t.forward(width)
        t.left(90)
        t.forward(height)
        t.left(90)
    t.end_fill()
    t.setheading(0)  # Resetear a dirección original
def draw_rectangle(t, width, height):
    global numeros_derecha
    global k
    x = coordinates[0]
    y = coordinates[1]

    t.penup()
    t.goto(x, y)
    t.pendown()

    # Establecer colores
    t.pencolor("gray")         # Color del contorno (plomo)

    t.fillcolor(color_palet[k])     # Color de relleno (celeste)
    # Marcar ocupado antes de avanzar
    key = (int(t.xcor())//10, int(t.ycor())//10)
    occupied_positions[key] = True
    t.begin_fill()
    # Dibujar y rellenar el rectángulo
    for _ in range(2):
        t.forward(width)
        t.left(90)
        t.forward(height)
        t.left(90)

    
    t.end_fill()
    # Actualizar coordenadas
    coordinates[0] += width
    coordinates[1] += height
    # numeros_derecha = np.append(numeros_derecha, [coordinates[0]+width , coordinates[1]+height])

LIMIT_X_MIN = -300
LIMIT_X_MAX = 300
LIMIT_Y_MIN = -300
LIMIT_Y_MAX = 300

def ramas(t, h, distancia_hasta_el_borde_y, pose_x, i, cantidad_de_ramas, count, nivel):
    global lados_rectangulos, secuencia, coordinates, numeros_derecha
    global okidoki

    if i <= 1 or count <= 0 or nivel > 5:
        return

    rama_height = int(secuencia[i - 1] / 2)
    rama_width = int(secuencia[i - 2] / 2)
    consumo = int((rama_height * rama_width) / 100)

    if count < consumo:
        return

    count -= consumo

    base_x = pose_x
    base_y = numeros_derecha[1] - distancia_hasta_el_borde_y

    # Evitar dibujar fuera de los límites
    if not (LIMIT_X_MIN <= base_x <= LIMIT_X_MAX and LIMIT_Y_MIN <= base_y <= LIMIT_Y_MAX):
        return

    lados_rectangulos = np.append(lados_rectangulos, rama_width)
    lados_rectangulos = np.append(lados_rectangulos, rama_height)

    okidoki = "okidoki" if base_y > -300 else "nothing"

    # Dibuja solo si está dentro de los límites
    if cantidad_de_ramas > 0:
        draw_same_squares_rectangle(rama_height * rama_width, 10, t, h, 600, 600, base_x)
        cantidad_de_ramas -= 1

    if cantidad_de_ramas > 0 and base_x + rama_width / 2 <= LIMIT_X_MAX:
        draw_same_squares_rectangle(rama_height * rama_width, 10, t, h, 600, 600, base_x + rama_width / 2)
        cantidad_de_ramas -= 1

    if cantidad_de_ramas > 0:
        draw_same_squares_rectangle(rama_height * rama_width, 10, t, h, 600, 600, base_x)
        cantidad_de_ramas -= 1

    if cantidad_de_ramas > 0 and base_x - rama_width / 2 >= LIMIT_X_MIN:
        draw_same_squares_rectangle(rama_height * rama_width, 10, t, h, 600, 600, base_x - rama_width / 2)
        cantidad_de_ramas -= 1

    i -= 2

    # Recursión segura solo si los nuevos puntos están dentro de los límites
    if base_y + rama_height <= LIMIT_Y_MAX:
        ramas(t, h + 2, distancia_hasta_el_borde_y + rama_height, pose_x, i, cantidad_de_ramas, count, nivel + 1)
    if base_y - rama_height >= LIMIT_Y_MIN:
        ramas(t, h + 2, distancia_hasta_el_borde_y - rama_height, pose_x, i, cantidad_de_ramas, count, nivel + 1)
    if base_x + rama_width <= LIMIT_X_MAX:
        ramas(t, h + 2, distancia_hasta_el_borde_y, pose_x + rama_width, i, cantidad_de_ramas, count, nivel + 1)
    if base_x - rama_width >= LIMIT_X_MIN:
        ramas(t, h + 2, distancia_hasta_el_borde_y, pose_x - rama_width, i, cantidad_de_ramas, count, nivel + 1)
def work_with_squares(t,rectangle_width,rectangle_height,distancia_hasta_el_borde_x,distancia_hasta_el_borde_y,old_width,old_heigth,h,distancia_hasta_el_borde_oldl):
    global perimetro
    global numeros_ramas_derecha
    global numeros_derecha
    global numeros_cuadrado
    global numeros_derecha_cuadrados
    global lados_ramas 
    global lados_rectangulos
    global secuencia
    global color_palet
    global k
    lado_cuadrado = 10    
    #print(f"lado_cuadrado: {lado_cuadrado}") 
    #print(f"rectangle_width: {rectangle_width}") 
    #print(f"rectangle_height: {rectangle_height}") 
    coordinates[0]=numeros_derecha[-2]-lados_rectangulos[h]+60#/2
    area_total=abs(rectangle_width*rectangle_height)
    count= int(area_total/lado_cuadrado**2)
    
    if(rectangle_height > distancia_hasta_el_borde_y or rectangle_width > 600) :
       #print("h:", h)
       #print("lados_rectangulos[h]:", lados_rectangulos[h])
       if(lados_rectangulos[h]>0):
        #print("El rectangulo no cabe en el espacio restante")  
        count= int(lados_rectangulos[h]*lados_rectangulos[h+1]/100)
        #print("count:", count)
       
        coordinates[0]=numeros_derecha[-2]-lados_rectangulos[h]#/2
        coordinates[1]=numeros_derecha[-1]
        numeros_derecha = np.append(numeros_derecha,coordinates)
        secuencia = fibonacci_dual(int(rectangle_width), int(distancia_hasta_el_borde_y))
        
        #print("Secuencia generada:", secuencia)
        #print("distancia_hasta_el_borde_y:", distancia_hasta_el_borde_y)
        #print("distancia_hasta_el_borde_x:", distancia_hasta_el_borde_x)
       #print("Longitud de la secuencia:", len(secuencia))  # Usamos len() en la lista resultante

     
        pose_x= -(600 / 2) 
        cantidad_de_ramas=int(((secuencia[-1]*secuencia[-2])/100)/((lados_rectangulos[h]*lados_rectangulos[h+1])/100))
        saved_h = h
        h=len(lados_rectangulos)
        # Agregar dimensiones de la primera rama al arreglo global
        lados_rectangulos = np.append(lados_rectangulos, lados_rectangulos[saved_h])
        lados_rectangulos = np.append(lados_rectangulos,  lados_rectangulos[saved_h+1])
        ramas(t,h,distancia_hasta_el_borde_y,pose_x,len(secuencia),cantidad_de_ramas,count,0)
        h = saved_h

        coordinates[0]=(numeros_derecha[0]/2)-(lados_rectangulos[h]/2)
        coordinates[1]=numeros_derecha[h]-distancia_hasta_el_borde_y
       #print(f"Perimetro ahora1: {perimetro}")
          
    if(rectangle_height <= distancia_hasta_el_borde_y and rectangle_width < 600) :
        
       if(lados_rectangulos[h]>0):  
         coordinates[0]=numeros_derecha[0]-lados_rectangulos[0]/2-lados_rectangulos[h]/2
        #print("El rectangulo cabe en el espacio restante") 
         new_heigth=rectangle_height
         new_width=rectangle_width
         perimetro+=(new_width*2)+(new_heigth*2)
        #print(f"Perimetro ahora: {perimetro}") 
         
        #print(f"lado_cuadrado: {lado_cuadrado}") 
        #print(f"coordinates: {coordinates}") 
        #print(f"t: {t}") 
        #print(f"h: {h}") 
         count= int(area_total/lado_cuadrado**2)
        #print(f"count: {count}")
        #print(f"coordinates[1] en work_with_squares: {coordinates[1]}") 
         if(count>0):
          draw_same_squares_rectangle(count,lado_cuadrado,t,h,old_width,old_heigth,numeros_derecha[0]-(600/2))
    if(k<5):
     k+=1
   #print(f"k: {k}")     
def draw_squares_around_the_fig(cantidad,lados_rectangulos,numeros_derecha,lado_cuadrado,coordinates,numeros_cuadrado,t,numeros_derecha_cuadrados,numeros,distancia_hasta_el_borde_old):
    espacio_del_cuadrado=int(perimetro/cantidad)
    direction_flag=True 
    coordinates_i=coordinates
    coordinates_j=coordinates
    i=0
    j=0
    
    old_width=600
    old_heigth=600
    
    espacio_del_cuadrado=espacio_del_cuadrado*lado_cuadrado*0.5
    while(j<old_width): #x
     
     if((j%espacio_del_cuadrado == 0) and (j >= espacio_del_cuadrado)):
            coordinates[0] =-old_width+j+(lado_cuadrado/2)
            coordinates[1] =-(old_width/2)

            
            numeros_cuadrado= np.append(numeros_cuadrado, coordinates)
            numeros_derecha_cuadrados=draw_rectangle(t,coordinates, lado_cuadrado, lado_cuadrado,numeros_derecha_cuadrados)  
            
            
     j+=1
          
    while(i<old_width):
        
        if((i%espacio_del_cuadrado == 0) and (i >= espacio_del_cuadrado)):
            coordinates[0]=-old_width+j-espacio_del_cuadrado
            coordinates[1]=(-old_heigth/2)+i+(lado_cuadrado/2)
    
            
            numeros_cuadrado= np.append(numeros_cuadrado, coordinates)
            numeros_derecha_cuadrados=draw_rectangle(t,coordinates, lado_cuadrado, lado_cuadrado,numeros_derecha_cuadrados)  

            
        i+=lado_cuadrado

    h=0  
    j=0
    i=0   

    coordinates[0]=numeros_derecha[2]-lados_rectangulos[0]
    coordinates[1]=numeros_derecha[3]-100
     
    direction_flag=True 
    draw_square(h,i,j,numeros_derecha[2]-lados_rectangulos[0],numeros_derecha[3]-lados_rectangulos[1],numeros_derecha[2]-lados_rectangulos[0],numeros_derecha[3],lado_cuadrado,numeros_derecha,lados_rectangulos,espacio_del_cuadrado,numeros_derecha_cuadrados,numeros_cuadrado,t,direction_flag)
    # Guardar la coordenada superior izquierda
    numeros_derecha = np.append(numeros_derecha, [coordinates[0]+lados_rectangulos[h] , coordinates[1]+lados_rectangulos[h+1]])
   
    #numeros derecha incluyen el sado superior del cuadrado anfitrion

    h=2
    i=0   
    j=0 
    
    while h < (lados_rectangulos.size-2 ):
       if numeros_derecha[1]-numeros_derecha[h+1] <= lados_rectangulos[h+1]:

        break
       direction_flag=True
       draw_square(h,i,j,numeros_derecha[h+2]-lados_rectangulos[h],numeros_derecha[h+3]-lados_rectangulos[h+1],numeros_derecha[h+2]-lados_rectangulos[h],numeros_derecha[h+3],lado_cuadrado,numeros_derecha,lados_rectangulos,espacio_del_cuadrado,numeros_derecha_cuadrados,numeros_cuadrado,t,direction_flag)
       numeros_derecha = np.append(numeros_derecha, [coordinates[0]+lados_rectangulos[h] , coordinates[1]+lados_rectangulos[h+1]])
       h+=2
       i=0   
       j=0  
       coordinates[0]=numeros_derecha[h]-(lados_rectangulos[h-4]/2)-(lados_rectangulos[h]/2) 
       coordinates[1]=numeros_derecha[h+1] 
       
   #print(f"Lado top: {lados_rectangulos[h-2]}")   
    h-=2
    
    j=lados_rectangulos[h]
    i=lados_rectangulos[h+1]
    coordinates[0]=numeros_derecha[h]-(lados_rectangulos[h-4]/2)-(lados_rectangulos[h]/2) 
    coordinates[1]=numeros_derecha[h+1] 
    direction_flag=False

    draw_square(h,i,j,numeros_derecha[h+2],numeros_derecha[h+3],numeros_derecha[h+2],numeros_derecha[h+3]-lados_rectangulos[h],lado_cuadrado,numeros_derecha,lados_rectangulos,espacio_del_cuadrado,numeros_derecha_cuadrados,numeros_cuadrado,t,direction_flag)
      
    h-=2
    
    coordinates[0]=numeros_derecha[h+2] 
    coordinates[1]=numeros_derecha[h+3] 

    
    while(h>=0): 
      direction_flag=False
      j=lados_rectangulos[h]
      i=lados_rectangulos[h+1]
      draw_square(h,i,j,numeros_derecha[h+2],numeros_derecha[h+3],numeros_derecha[h+2],numeros_derecha[h+3]-lados_rectangulos[h],lado_cuadrado,numeros_derecha,lados_rectangulos,espacio_del_cuadrado,numeros_derecha_cuadrados,numeros_cuadrado,t,direction_flag) 
   
      h-=2
    i=0
    j=0  
     
    coordinates[0]=numeros_derecha[0]-old_width
    coordinates[1]=numeros_derecha[1] 
def draw_main_rectangle(t, width, height):
    global numeros_derecha
    global k
    global coordinates
    x = coordinates[0]
    y = coordinates[1]


    t.penup()
    t.goto(x, y)
    t.pendown()

    t.pencolor("gray")
    t.fillcolor(color_palet[k])

    t.begin_fill()
    for _ in range(2):
        t.forward(width)
        t.left(90)
        t.forward(height)
        t.left(90)
    t.end_fill()



    # Actualizar coordenadas para el siguiente rectángulo en la misma fila
    coordinates[0] += width
    coordinates[1] += height          
def draw_nested_rectangles():
    global numeros
    global numeros_derecha
    global perimetro
    distancia_hasta_el_borde=0
    distancia_hasta_el_borde_old=0
    perimetro = 0
    global coordinates
    modo_reducido = False
    numeros = np.array([])
    
    global lados_rectangulos
    global color_palet
    lados_rectangulos = np.append(lados_rectangulos, [210,50])
    lados_rectangulos = np.append(lados_rectangulos, [150, 180])
    lados_rectangulos = np.append(lados_rectangulos, [40, 500])
    lados_rectangulos = np.append(lados_rectangulos, [50, 100])
    lados_rectangulos = np.append(lados_rectangulos, [0, 0]) #finaliza la sequencia   
    
    color_palet = np.append(color_palet, "white")
    #===========================================
    color_palet = np.append(color_palet, "red")
    color_palet = np.append(color_palet, "orange")
    color_palet = np.append(color_palet, "yellow")
    color_palet = np.append(color_palet, "green")
    color_palet = np.append(color_palet,"blue") #finaliza la sequencia   

   #print(f"lados_rectangulos: {lados_rectangulos}") 
    global lados_cuadrado
    global numeros_cuadrado
    global numeros_cuadrado_derecha
    global k
    lado_cuadrado = 10
    window = turtle.Screen()
    window.title("División de Rectángulos")
    window.setup(width=600, height=600)  
    window.tracer(0)  

    t = turtle.Turtle()
    t.speed(0)  

    # Dibujar el rectángulo principal y actualizar coordenadas
    main_width = window.window_width()
    main_height = window.window_height()
    coordinates[0] -=main_width
   
    coordinates[1] -=(main_height/2)
   #print(f"coordinates: {coordinates}")  
    #cuadrado anfitrion

    numeros=np.append(numeros,coordinates)
   #print(f"numeros_derecha: {numeros_derecha}")    
    draw_main_rectangle(t, main_width, main_height)
    
    k+=1
   #print(f"k: {k}") 
    numeros_derecha = np.append(numeros_derecha, [coordinates[0] , coordinates[1]])
    print(f"coordinates[0] {coordinates[0]} numeros_derecha[0]={numeros_derecha[0]}")
    print(f"coordinates[1] {coordinates[1]} numeros_derecha[1]={numeros_derecha[1]}")
   #print(f"numeros_derecha: {numeros_derecha}") 

    perimetro+=(2*main_width)+(2*main_height)


    
    coordinates[0]=coordinates[0]+(main_width/2)-(lados_rectangulos[0]/2)
    coordinates[1]=coordinates[1]-main_height
    numeros_derecha = np.append(numeros_derecha, coordinates) 
   #print(f"coordinates: {coordinates}")  
    
    h=0
    
    while(h<(lados_rectangulos.size-2)):
     # Dibujar segundo rectángulo sobre el primero

    #print(f"h: {h}")      
    #print(f"lados_rectangulos[h]: {lados_rectangulos[h]}") 
    #print(f"lados_rectangulos[h+1]: {lados_rectangulos[h+1]}") 
      
     distancia_hasta_el_borde_y = abs(numeros_derecha[1]-coordinates[1])
     distancia_hasta_el_borde_x = (main_width/2) 
    
    #print(f"distancia_hasta_el_borde_y : {distancia_hasta_el_borde_y }") 
    #print(f"distancia_hasta_el_borde_x: {distancia_hasta_el_borde_x}")
   

     work_with_squares(t,lados_rectangulos[h],lados_rectangulos[h+1],distancia_hasta_el_borde_x,distancia_hasta_el_borde_y,main_width,main_height,h,distancia_hasta_el_borde_old)     
                  
  
     h+=2
     coordinates[0]=lados_rectangulos[h-2]-(lados_rectangulos[h-2]/2)-(lados_rectangulos[h]/2) #coordinates[0]=numeros_derecha[h]-(lados_rectangulos[h-2]/2)-(main_width/2)
     
    #print(f"coordinates[1] en nested rectangles: {coordinates[1]}")
    #print(f"coordinates[1]: {coordinates[1]}")
     numeros_derecha = np.append(numeros_derecha, coordinates) 

    # Finalizar
    window.update()
    window.mainloop()   
# Parámetros: rectángulo principal y subrectángulo

draw_nested_rectangles()