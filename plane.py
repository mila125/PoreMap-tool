from mpl_toolkits.mplot3d import Axes3D

# Crear datos sintéticos de microporosidad
x, y = np.meshgrid(range(100), range(100))
z = np.random.randint(0, 2, size=(100, 100))  # Simulación de porosidad

# Visualización 3D
fig = plt.figure(figsize=(8, 6))
ax = fig.add_subplot(111, projection='3d')
ax.plot_surface(x, y, z, cmap='viridis')

ax.set_title("Visualización 3D de Microporosidad")
plt.show()