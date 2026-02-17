import pygame
import random
import sys

# =============================
# Configuración e inicialización
# =============================
pygame.init()
pygame.font.init()

ANCHO = 1000
ALTO = 650
FPS = 60

pantalla = pygame.display.set_mode((ANCHO, ALTO))
pygame.display.set_caption("Sunatsitos ¡Bienvenidos a SUNAT")

# Fuentes
FUENTE = pygame.font.SysFont("Arial", 30)
FUENTE_GRANDE = pygame.font.SysFont("Arial", 45)

# Colores
BLANCO = (255, 255, 255)
AZUL = (0, 102, 204)
AMARILLO = (255, 204, 0)
VERDE = (0, 180, 0)
ROJO = (200, 0, 0)
NEGRO = (0, 0, 0)
GRIS = (230, 230, 230)

# =============================
# Parámetros de juego
# =============================
PLAYER_SPEED = 9

# Velocidades "amistosas" para 6 años (px por frame a ~60 FPS)
BASE_FALL_SPEED_EASY = 3.0   # ~240 px/s  → tarda ~2.7s en caer
BASE_FALL_SPEED_MED  = 5.5   # ~390 px/s  → ~1.7s
BASE_FALL_SPEED_HARD = 8.0   # ~540 px/s  → ~1.2s
BASE_FALL_SPEED_MAX3 = 10.0  # meta del final del minuto 3 (rápido pero jugable)

# Spawns por minuto (menor número => más aparición) Easy, Medium, Hard
SPAWN_FREQ_VERDE = [22, 18, 14]
SPAWN_FREQ_ROJO  = [32, 26, 20]

ANCHO_OBJ = 40
ALTO_OBJ  = 40

TOTAL_DUR_MS = 180_000          # 3 minutos
TRIVIA_INTERVAL_MS = 60_000     # cada minuto: 60s y 120s
TRIVIA_PREGUNTAS_POR_EVENTO = 3
TRIVIA_PUNTOS_POR_ACIERTO = 10
FEEDBACK_MS = 800               # feedback de color en trivia
INVULN_AFTER_TRIVIA_MS = 1200   # invulnerable al volver de trivia

# Fondos por minuto: fácil, medio, difícil
BG_POR_NIVEL = [
    (220, 240, 255),  # fácil
    (235, 240, 220),  # medio
    (255, 230, 230),  # difícil
]

# Bloques amarillos por minuto (fácil, medio, difícil)
YELLOWS_PER_MINUTE = [1, 2, 3]

# =============================
# Utilidades
# =============================
def texto(msg, fuente, color, x, y):
    t = fuente.render(msg, True, color)
    pantalla.blit(t, (x, y))

def boton(msg, x, y, w, h, color):
    rect = pygame.Rect(x, y, w, h)
    pygame.draw.rect(pantalla, color, rect, border_radius=12)
    # Texto: negro sobre amarillo, blanco sobre otros colores
    t = FUENTE.render(msg, True, NEGRO if color == AMARILLO else BLANCO)
    tx = x + (w - t.get_width()) // 2
    ty = y + (h - t.get_height()) // 2
    pantalla.blit(t, (tx, ty))
    return rect

def formato_tiempo(ms):
    if ms < 0: ms = 0
    total_seg = ms // 1000
    m = total_seg // 60
    s = total_seg % 60
    return f"{m:02d}:{s:02d}"

def dibujar_vidas(vidas_restantes, x_der=ANCHO-20, y=25, radio=9, separacion=28):
    """Dibuja 'vidas_restantes' puntos rojos en la esquina superior derecha."""
    for i in range(vidas_restantes):
        cx = x_der - i * separacion
        pygame.draw.circle(pantalla, ROJO, (cx, y), radio)

def lerp(a, b, t):
    return a + (b - a) * max(0.0, min(1.0, t))

# =============================
# REGISTRO
# =============================
def registro():
    nombre = ""
    genero = ""
    visita = ""
    reloj = pygame.time.Clock()
    cursor_visible = True
    cursor_timer = 0

    while True:
        pantalla.fill(AZUL)
        # Título
        texto("Sunatsitos ¡Bienvenidos a SUNAT", FUENTE_GRANDE, BLANCO, 80, 50)
        texto("Escribe tu nombre:", FUENTE, BLANCO, 100, 160)

        # Caja de texto
        caja = pygame.Rect(100, 200, 400, 45)
        pygame.draw.rect(pantalla, BLANCO, caja, border_radius=10)

        # Cursor parpadeante
        cursor_timer += reloj.get_time()
        if cursor_timer >= 500:
            cursor_visible = not cursor_visible
            cursor_timer = 0

        nombre_mostrar = nombre + ("|" if cursor_visible else "")
        texto(nombre_mostrar, FUENTE, NEGRO, 110, 210)

        # Botones
        b_nino = boton("Niño", 100, 300, 150, 60, VERDE)
        b_nina = boton("Niña", 270, 300, 150, 60, VERDE)
        b_papa = boton("Ver a Papá", 100, 380, 200, 60, AMARILLO)
        b_mama = boton("Ver a Mamá", 320, 380, 200, 60, AMARILLO)

        # Indicadores
        if genero:
            texto(f"Género: {genero}", FUENTE, BLANCO, 550, 315)
        if visita:
            texto(f"Visita: {visita}", FUENTE, BLANCO, 550, 395)

        b_cont = None
        if nombre.strip() and genero and visita:
            b_cont = boton("EMPEZAR", 100, 480, 220, 65, ROJO)

        for e in pygame.event.get():
            if e.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if e.type == pygame.KEYDOWN:
                if e.key == pygame.K_BACKSPACE:
                    nombre = nombre[:-1]
                elif e.key == pygame.K_RETURN and b_cont:
                    return nombre.strip(), genero, visita
                else:
                    if len(nombre) < 20 and e.unicode.isprintable():
                        nombre += e.unicode

            if e.type == pygame.MOUSEBUTTONDOWN:
                if b_nino.collidepoint(e.pos): genero = "Niño"
                if b_nina.collidepoint(e.pos): genero = "Niña"
                if b_papa.collidepoint(e.pos): visita = "Papá"
                if b_mama.collidepoint(e.pos): visita = "Mamá"
                if b_cont and b_cont.collidepoint(e.pos):
                    return nombre.strip(), genero, visita

        pygame.display.update()
        reloj.tick(FPS)

# =============================
# INSTRUCCIONES
# =============================
def instrucciones():
    while True:
        pantalla.fill(GRIS)
        texto("INSTRUCCIONES", FUENTE_GRANDE, AZUL, 360, 40)

        y = 120
        salt = 40
        indicaciones = [
            "Objetivo: Junta VERDES (dinero) para sumar puntos.",
            "Cuidado: Evita ROJOS (riesgo). ¡Te quitan una vida!",
            "Vida Extra: AMARILLO te da +1 vida.",
            "Teclas: Flechas Izquierda/Derecha para moverte.",
            "Tiempo total: 3 minutos con progresión suave:",
            "  - 0:00–1:00 Lento (Fácil, 1 amarillo)",
            "  - 1:00–2:00 Más rápido (Medio, 2 amarillos)",
            "  - 2:00–3:00 Mucho más rápido (Difícil, 3 amarillos)",
            "La velocidad aumenta suavemente todo el tiempo (sin saltos).",
            "Trivia: Aparece cada minuto (pausa el juego).",
            "En Trivia: si fallas → ROJO y pierdes 1 vida.",
            "Vidas: Puntos rojos arriba a la derecha.",
        ]
        for linea in indicaciones:
            texto(f"• {linea}", FUENTE, NEGRO, 80, y)
            y += salt

        # Botón movido a la mitad derecha para no tapar el texto
        b_jugar = boton("COMENZAR", 700, 520, 240, 65, AZUL)

        for e in pygame.event.get():
            if e.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if e.type == pygame.MOUSEBUTTONDOWN and b_jugar.collidepoint(e.pos):
                return

        pygame.display.update()

# =============================
# TRIVIA (3 preguntas, pausa el juego)
# =============================
def trivia_evento(puntos, vidas):
    preguntas = [
        ("¿Los impuestos NO ayudan a construir colegios?", False),
        ("¿El IGV se aplica a la venta de bienes y servicios?", True),
        ("¿Pedir boleta o factura NO ayuda a la formalidad?", False),
    ]
    idx = 0
    reloj = pygame.time.Clock()

    while idx < TRIVIA_PREGUNTAS_POR_EVENTO and idx < len(preguntas):
        pregunta, es_verdadero = preguntas[idx]

        # Dibujo de pantalla de trivia
        pantalla.fill(GRIS)
        texto("TRIVIA SUNAT", FUENTE_GRANDE, AZUL, 320, 70)
        texto(f"Pregunta {idx+1} de {TRIVIA_PREGUNTAS_POR_EVENTO}", FUENTE, NEGRO, 360, 140)
        texto(pregunta, FUENTE, NEGRO, 120, 220)

        # HUD
        texto(f"Puntos: {puntos}", FUENTE, ROJO, 20, 20)
        dibujar_vidas(vidas)

        # Botones V/F amarillos
        b_v = boton("VERDADERO", 300, 340, 180, 70, AMARILLO)
        b_f = boton("FALSO",     520, 340, 180, 70, AMARILLO)
        pygame.display.update()

        respuesta_dada = False
        correcta = False
        boton_presionado = None

        # Esperar click
        while not respuesta_dada:
            for e in pygame.event.get():
                if e.type == pygame.QUIT:
                    pygame.quit()
                    sys.exit()
                if e.type == pygame.MOUSEBUTTONDOWN:
                    if b_v.collidepoint(e.pos):
                        boton_presionado = "V"
                        correcta = es_verdadero is True
                        respuesta_dada = True
                    elif b_f.collidepoint(e.pos):
                        boton_presionado = "F"
                        correcta = es_verdadero is False
                        respuesta_dada = True
            reloj.tick(FPS)

        # Feedback en color
        color_v = AMARILLO
        color_f = AMARILLO
        if boton_presionado == "V":
            color_v = VERDE if correcta else ROJO
        elif boton_presionado == "F":
            color_f = VERDE if correcta else ROJO

        if correcta:
            puntos += TRIVIA_PUNTOS_POR_ACIERTO
        else:
            vidas -= 1
            if vidas <= 0:
                # Mostrar el último estado
                pantalla.fill(GRIS)
                texto("TRIVIA SUNAT", FUENTE_GRANDE, AZUL, 320, 70)
                texto(pregunta, FUENTE, NEGRO, 120, 220)
                texto(f"Puntos: {puntos}", FUENTE, ROJO, 20, 20)
                dibujar_vidas(0)
                boton("VERDADERO", 300, 340, 180, 70, color_v)
                boton("FALSO",     520, 340, 180, 70, color_f)
                pygame.display.update()
                pygame.time.delay(FEEDBACK_MS)
                return puntos, vidas

        # Redibujar con feedback visual
        pantalla.fill(GRIS)
        texto("TRIVIA SUNAT", FUENTE_GRANDE, AZUL, 320, 70)
        texto(f"Pregunta {idx+1} de {TRIVIA_PREGUNTAS_POR_EVENTO}", FUENTE, NEGRO, 360, 140)
        texto(pregunta, FUENTE, NEGRO, 120, 220)
        texto(f"Puntos: {puntos}", FUENTE, ROJO, 20, 20)
        dibujar_vidas(vidas)
        boton("VERDADERO", 300, 340, 180, 70, color_v)
        boton("FALSO",     520, 340, 180, 70, color_f)
        pygame.display.update()
        pygame.time.delay(FEEDBACK_MS)

        idx += 1

    return puntos, vidas

# =============================
# Curva de velocidad PROGRESIVA (fácil -> medio -> difícil)
# =============================
def calcular_fall_speed(elapsed_ms):
    """
    Progresión suave durante 3 minutos:
    - 0–60s:     EASY  -> MED    (incremento continuo)
    - 60–120s:   MED   -> HARD   (incremento continuo)
    - 120–180s:  HARD  -> MAX3   (incremento continuo; más rápido pero jugable)
    """
    if elapsed_ms < 60_000:
        t = elapsed_ms / 60_000.0
        return lerp(BASE_FALL_SPEED_EASY, BASE_FALL_SPEED_MED, t)
    elif elapsed_ms < 120_000:
        t = (elapsed_ms - 60_000) / 60_000.0
        return lerp(BASE_FALL_SPEED_MED, BASE_FALL_SPEED_HARD, t)
    else:
        t = (elapsed_ms - 120_000) / 60_000.0
        return lerp(BASE_FALL_SPEED_HARD, BASE_FALL_SPEED_MAX3, t)

# =============================
# JUEGO PRINCIPAL (3 minutos, niveles por minuto)
# =============================
def juego(nombre, visita):
    jugador = pygame.Rect(470, 540, 70, 70)
    objetos_verdes = []
    objetos_rojos = []
    powerups_amarillos = []

    puntos = 0
    vidas = 3

    reloj = pygame.time.Clock()
    elapsed_total_ms = 0
    minute_start_ms = 0
    next_trivia_ms = TRIVIA_INTERVAL_MS  # 60s, luego 120s
    yellows_spawned_this_min = 0
    invuln_ms = 0  # invulnerabilidad al volver de trivia

    running = True
    while running:
        dt = reloj.tick(FPS)  # ms desde el último frame
        elapsed_total_ms += dt
        invuln_ms = max(0, invuln_ms - dt)

        # Fin por tiempo
        if elapsed_total_ms >= TOTAL_DUR_MS:
            return puntos, "Tiempo cumplido"

        # Índice de minuto y parámetros
        minute_idx = min(elapsed_total_ms // 60_000, 2)  # 0=Fácil, 1=Medio, 2=Difícil
        # Detectar cambio de minuto
        if elapsed_total_ms - minute_start_ms >= 60_000 or (elapsed_total_ms < 60_000 and minute_start_ms != 0):
            # iniciar nuevo minuto
            minute_start_ms = (elapsed_total_ms // 60_000) * 60_000
            yellows_spawned_this_min = 0  # reset de amarillos por minuto

        # Velocidad de caída progresiva
        fall_speed = calcular_fall_speed(elapsed_total_ms)

        # Eventos básicos
        for e in pygame.event.get():
            if e.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

        # Movimiento del jugador
        teclas = pygame.key.get_pressed()
        if teclas[pygame.K_LEFT]:
            jugador.x -= PLAYER_SPEED
        if teclas[pygame.K_RIGHT]:
            jugador.x += PLAYER_SPEED
        jugador.x = max(0, min(jugador.x, ANCHO - jugador.width))

        # Spawns (probabilísticos por frame)
        if random.randint(1, SPAWN_FREQ_VERDE[int(minute_idx)]) == 1:
            objetos_verdes.append(
                pygame.Rect(random.randint(0, ANCHO - ANCHO_OBJ), 0, ANCHO_OBJ, ALTO_OBJ)
            )
        if random.randint(1, SPAWN_FREQ_ROJO[int(minute_idx)]) == 1:
            objetos_rojos.append(
                pygame.Rect(random.randint(0, ANCHO - ANCHO_OBJ), 0, ANCHO_OBJ, ALTO_OBJ)
            )

        # Spawns de amarillos (máximo por minuto)
        objetivo_yellow = YELLOWS_PER_MINUTE[int(minute_idx)]
        if yellows_spawned_this_min < objetivo_yellow:
            # baja probabilidad por frame hasta completar el cupo del minuto
            if random.randint(1, 240) == 1:  # ~0.4% por frame
                powerups_amarillos.append(
                    pygame.Rect(random.randint(0, ANCHO - ANCHO_OBJ), 0, ANCHO_OBJ, ALTO_OBJ)
                )
                yellows_spawned_this_min += 1

        # Actualizar objetos verdes
        for obj in objetos_verdes[:]:
            obj.y += fall_speed
            if jugador.colliderect(obj):
                puntos += 1
                objetos_verdes.remove(obj)
            elif obj.y > ALTO:
                objetos_verdes.remove(obj)

        # Actualizar objetos rojos
        for obj in objetos_rojos[:]:
            obj.y += fall_speed
            if jugador.colliderect(obj) and invuln_ms <= 0:
                vidas -= 1
                objetos_rojos.remove(obj)
                if vidas <= 0:
                    return puntos, "Sin vidas"
            elif obj.y > ALTO:
                objetos_rojos.remove(obj)

        # Actualizar power-ups amarillos
        for obj in powerups_amarillos[:]:
            obj.y += fall_speed
            if jugador.colliderect(obj):
                vidas += 1
                powerups_amarillos.remove(obj)
            elif obj.y > ALTO:
                powerups_amarillos.remove(obj)

        # Lanzar trivia en los minutos (60s, 120s) — pausa del juego
        if elapsed_total_ms >= next_trivia_ms and next_trivia_ms < TOTAL_DUR_MS:
            puntos, vidas = trivia_evento(puntos, vidas)
            invuln_ms = INVULN_AFTER_TRIVIA_MS
            next_trivia_ms += TRIVIA_INTERVAL_MS
            if vidas <= 0:
                return puntos, "Sin vidas"

        # Dibujo del frame (fondo según minuto: fácil, medio, difícil)
        pantalla.fill(BG_POR_NIVEL[int(minute_idx)])

        # Jugador
        pygame.draw.rect(pantalla, AZUL, jugador, border_radius=6)

        # Objetos
        for obj in objetos_verdes:
            pygame.draw.rect(pantalla, VERDE, obj, border_radius=6)
        for obj in objetos_rojos:
            pygame.draw.rect(pantalla, ROJO, obj, border_radius=6)
        for obj in powerups_amarillos:
            pygame.draw.rect(pantalla, AMARILLO, obj, border_radius=6)

        # HUD
        tiempo_restante = TOTAL_DUR_MS - elapsed_total_ms
        texto(f"Puntos: {puntos}", FUENTE, ROJO, 20, 20)
        texto(f"Jugador: {nombre}", FUENTE, NEGRO, 20, 55)
        texto(f"Visita: {visita}", FUENTE, NEGRO, 20, 90)
        texto(f"Tiempo: {formato_tiempo(tiempo_restante)}", FUENTE, NEGRO, 20, 125)
        # Etiqueta de nivel coherente con la velocidad
        nivel_txt = ["Fácil", "Medio", "Difícil"][int(minute_idx)]
        texto(f"Nivel: {nivel_txt}", FUENTE, NEGRO, 20, 160)

        dibujar_vidas(vidas)

        pygame.display.update()

# =============================
# FINAL (con reinicio sin pedir nombre)
# =============================
def final(nombre, visita, puntos, motivo):
    while True:
        pantalla.fill(AZUL)
        texto("¡FIN DEL JUEGO!", FUENTE_GRANDE, BLANCO, 360, 120)
        texto(nombre, FUENTE_GRANDE, AMARILLO, 420, 190)
        texto(f"Visitaste SUNAT para ver a tu {visita}", FUENTE, BLANCO, 260, 260)
        texto(f"Puntaje final: {puntos}", FUENTE_GRANDE, BLANCO, 360, 330)

        if motivo == "Sin vidas":
            texto("Motivo: Te quedaste sin vidas 💥", FUENTE, BLANCO, 340, 390)
        elif motivo == "Tiempo cumplido":
            texto("¡Tiempo cumplido! ⏱️", FUENTE, BLANCO, 390, 390)

        # Botones: Jugar de nuevo / Salir
        b_reiniciar = boton("JUGAR DE NUEVO", 300, 470, 220, 65, VERDE)
        b_salir     = boton("SALIR",          540, 470, 160, 65, ROJO)

        for e in pygame.event.get():
            if e.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if e.type == pygame.MOUSEBUTTONDOWN:
                if b_reiniciar.collidepoint(e.pos):
                    return "reiniciar"
                if b_salir.collidepoint(e.pos):
                    return "salir"

        pygame.display.update()

# =============================
# MAIN (bucle de reinicio sin pedir nombre)
# =============================
def main():
    nombre, genero, visita = registro()
    while True:
        instrucciones()
        puntos, motivo = juego(nombre, visita)
        accion = final(nombre, visita, puntos, motivo)
        if accion == "reiniciar":
            # reinicia solo la partida, conservando nombre/género/visita
            continue
        else:
            break

if __name__ == "__main__":
    main()
