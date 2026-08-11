import unittest
from app import app # Importamos tu sistema real

class PruebasCajaBlancaSeguridad(unittest.TestCase):
    
    def setUp(self):
        # Esta función prepara el "clon" de tu sistema antes de cada prueba
        app.config['TESTING'] = True
        self.tester = app.test_client()

    def test_bloqueo_intrusos_sin_sesion(self):
        # PRUEBA 1: Intentamos entrar a /clientes SIN haber iniciado sesión
        respuesta = self.tester.get('/clientes')
        
        # Evaluamos que el sistema internamente nos responda con un código 302 (Redirección)
        self.assertEqual(respuesta.status_code, 302)
        # Evaluamos que la redirección nos mande obligatoriamente al '/login'
        self.assertTrue('/login' in respuesta.headers['Location'])

    def test_bloqueo_metodo_incorrecto_api(self):
        # PRUEBA 2: Intentamos hacerle un GET a la ruta de consultar documento 
        # (cuando sabemos que está programada estrictamente para POST)
        respuesta = self.tester.get('/api/consultar_documento')
        
        # Evaluamos que el sistema lo rechace con código 405 (Method Not Allowed)
        self.assertEqual(respuesta.status_code, 405)

if __name__ == '__main__':
    unittest.main()