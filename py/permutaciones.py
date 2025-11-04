def generar_permutaciones(elementos):
    if len(elementos) == 0:
        return [[]]

    primera_elemento = elementos[0]
    restantes_elementos = elementos[1:]

    # Generar permutaciones recursivamente para los elementos restantes
    permutaciones_restantes = generar_permutaciones(restantes_elementos)

    # Insertar el primer elemento en todas las posiciones de cada permutación restante
    todas_permutaciones = []
    for perm in permutaciones_restantes:
        for i in range(len(perm) + 1):
            nueva_permutacion = perm[:i] + [primera_elemento] + perm[i:]
            todas_permutaciones.append(nueva_permutacion)

    return todas_permutaciones

# Ejemplo: Mostrar las permutaciones de 3 elementos
elementos = [ '1','2','3','4','5','6','7','8']
todas_permutaciones = generar_permutaciones(elementos)

for perm in todas_permutaciones:
    print(perm)