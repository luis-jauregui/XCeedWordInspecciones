namespace XCeedWordInspeccion
{
    public class Model
    {

        public class Ensayo
        {
           public string IdProducto { get; set; }
           public string NombreGenerico { get; set; }
           public int IdAnalisis { get; set; }
           public int IdMetodo { get; set; }
           public string Analisis { get; set; }
           public string Metodo { get; set; }
           public string UnidadMedida { get; set; }
           public string ResultadoDef { get; set; }
           public string LimiteC { get; set; }
           public string LimiteD { get; set; }
           public string Abreviatura { get; set; }
        }
        
        public class LoteCls
        {
            public string Lote { get; set; }
        }
        
        public class CodigoVia
        {
            public int IdProducto { get; set; }
            public string ProductoCodigo { get; set; }
            public string Vias { get; set; }
            public string CodigoInterno { get; set; }
        }
        
        public class CodigoViaFisicoSensorial
        {
            public int IdProducto { get; set; }
            public string CodigoInterno { get; set; }
            public string RangoVias { get; set; }
        }

        public class Via
        {
            public int NroViaTemporal;
            public int NroVia;
            public string Presentacion;
            public string PresentacionMuestra;
            public string Muestra;
            public int NumeroMuestra;
            public int IdProducto;
            public string Producto;
            public bool EsAguaPotable;
            public bool EsAguaManantial;
            public string ViaData;
        }
        
        public class ViaResultado
        {
            public string IdProducto { get; set; }
            public int IdAnalisis { get; set; }
            public string UnidMedida { get; set; }
            public string Resultado { get; set; }
            public string CodPrecinto { get; set; }
            public string Muestra { get; set; }
            public string CodigoInterno { get; set; }
            public int NroVia { get; set; }
        }
        
        public class MuestraCls
        {
            public int NroViaTemporal;
            public int NroVia;
            public string Presentacion;
            public string PresentacionMuestra;
            public string Muestra;
            public int NumeroMuestra;
            public int IdProducto;
            public string Producto;
            public bool EsAguaPotable;
            public bool EsAguaManantial;
            public string ViaData;
        }

        public class SpGetTablaEvaluacionDobleCierre
        {
            public int IdProducto { get; set; }
            public string Codigos { get; set; }
            public long Vias { get; set; } // ROW_NUMBER() devuelve un tipo bigint, que se mapea a long en C#
            public string Latas { get; set; }

            public string Compacidad1 { get; set; }
            public string Compacidad2 { get; set; }
            public string Compacidad3 { get; set; }
            public string Compacidad4 { get; set; }
            public string Compacidad5 { get; set; }
            public string Compacidad6 { get; set; }
            public string Compacidad7 { get; set; }
            public string Compacidad8 { get; set; }

            public string Traslape1 { get; set; }
            public string Traslape2 { get; set; }
            public string Traslape3 { get; set; }
            public string Traslape4 { get; set; }
            public string Traslape5 { get; set; }
            public string Traslape6 { get; set; }
            public string Traslape7 { get; set; }
            public string Traslape8 { get; set; }

            public string Traslapem1 { get; set; }
            public string Traslapem2 { get; set; }
            public string Traslapem3 { get; set; }
            public string Traslapem4 { get; set; }
            public string Traslapem5 { get; set; }
            public string Traslapem6 { get; set; }
            public string Traslapem7 { get; set; }
            public string Traslapem8 { get; set; }

            public string PenetracionDeGanchoDeCuerpo1 { get; set; }
            public string PenetracionDeGanchoDeCuerpo2 { get; set; }
            public string PenetracionDeGanchoDeCuerpo3 { get; set; }
            public string PenetracionDeGanchoDeCuerpo4 { get; set; }
            public string PenetracionDeGanchoDeCuerpo5 { get; set; }
            public string PenetracionDeGanchoDeCuerpo6 { get; set; }
            public string PenetracionDeGanchoDeCuerpo7 { get; set; }
            public string PenetracionDeGanchoDeCuerpo8 { get; set; }

            public string Arrugas { get; set; }
            public string DefectosVisibles { get; set; }
        }
        
        public class UspGetTablaExamenesSensoriales
        {
            public int IdProducto { get; set; }
            public string Codigos { get; set; }
            public long Vias { get; set; }
            public string RangoVias { get; set; }
            public string Latas { get; set; }
            public string EnvaseInterno { get; set; }
            public string EnvaseExterno { get; set; }
            public string VacioEntreEnvase { get; set; }
            public string PresionDeVacioMmHg { get; set; }
            public decimal Bruto { get; set; }
            public decimal SinLiquido { get; set; }
            public decimal Tara { get; set; }
            public decimal Escurrido { get; set; }
            public decimal Neto { get; set; }
            public string Contenido { get; set; }
            public string CondicionLiquidoGobierno { get; set; }
            public string Contenido3 { get; set; }
            public string Olor { get; set; }
            public string Color { get; set; }
            public string Sabor { get; set; }
            public string Textura { get; set; }
            public string LiquidoLibre { get; set; }
            public string PresenciaSal { get; set; }
        }
        
        public class UspGetReporteInspeccionTablaHistamina
        {
            public string IdProducto { get; set; }
            public int IdAnalisis { get; set; }
            public string UnidMedida { get; set; }
            public string Resultado { get; set; }
            public int NroVia { get; set; }
            public string CodPrecinto { get; set; }
            public string Muestra { get; set; }
            public string CodigoInterno { get; set; }
        }

        public class UspGetMuestraLaboratorioDirimente
        {
            public int IdCotizacion { get; set; }
            public int IdTipoAnalisis { get; set; }
            public string CodInterno { get; set; }
            public string MuestraLaboratorio { get; set; }
            public string MuestraDirimente { get; set; }
            // Exclusivo para FS
            public string MuestraLaboratorioCierre { get; set; }
            public string MuestraDirimenteCierre { get; set; }
        }

        public class UspGetListarAnalisisCNNuevo
        {
            public int IdProducto { get; set; }
            public string Codigos { get; set; }
            public int NVia  { get; set; }
            public string Latas { get; set; }
            public string Aspecto { get; set; }
            public string Color { get; set; }
            public string Olor { get; set; }
            public string Sabor { get; set; }
            public string Textura { get; set; }
            public string Liquido { get; set; }
        }
        
        public class FisicoSensorialCongeladoExportacionTodoDestino
        {
            public int IdProducto { get; set; }
            public string Codigos { get; set; }
            public int NVia  { get; set; }
            public string ProductoCodigo { get; set; }
            public string Latas { get; set; }
            public string Aspecto { get; set; }
            public string Color { get; set; }
            public string Olor { get; set; }
            public string Sabor { get; set; }
            public string Textura { get; set; }
            public string Liquido { get; set; }
        }
        
        
        public class Analisis
        {
            public int Id { get; set; }
            public string Label { get; set; }
        }
        
    }
}