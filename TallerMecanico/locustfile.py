from locust import HttpUser, task, between

class TallerMecanicoTester(HttpUser):
    # Tiempo de espera simulado entre clics (como un humano real leyendo la pantalla)
    wait_time = between(1, 3)

    @task
    def cargar_pantalla_login(self):
        # El bot entrará a la ruta /login repetidas veces
        self.client.get("/login")
        
    # Nota para tu tesis: Para hacer pruebas más avanzadas, los bots 
    # también podrían simular iniciar sesión enviando un POST con usuario y contraseña.
    # Por ahora probaremos la resistencia bruta de carga de vistas.