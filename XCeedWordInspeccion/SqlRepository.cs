using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Dapper;

namespace XCeedWordInspeccion
{
    public class SqlRepository
    {
        private readonly string _connectionString;
        
        public SqlRepository()
        {
            _connectionString = "Server=10.10.3.4;Database=certens_db;User Id=UserDB;Password=General480;";
        }

        public IEnumerable<T> ObtenerEnsayos<T>(int idOt, int idLaboratorio, int correlativo)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string storedProcedure = "up_get_VerAnalisiMetPorLabProd1"; // Nombre del procedimiento almacenado
        
                return connection.Query<T>(
                    storedProcedure,
                    new { IdOT = idOt, IdLaboratorio = idLaboratorio, Correlativo = correlativo },
                    commandType: CommandType.StoredProcedure // Indicar que es un SP
                );
            }
        }
        
        public IEnumerable<T> ObtenerCodigoVias<T>(string numOs)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string storedProcedure = "usp_get_ProductosPorNumOs"; // Nombre del procedimiento almacenado
        
                return connection.Query<T>(
                    storedProcedure,
                    new { NumOs = numOs },
                    commandType: CommandType.StoredProcedure // Indicar que es un SP
                );
            }
        }
        
        public IEnumerable<T> ObtenerVias<T>(int idOt, int correlativo, int idLaboratorio)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string storedProcedure = "usp_get_CodigoInternoViasPorIdOt"; // Nombre del procedimiento almacenado
        
                return connection.Query<T>(
                    storedProcedure,
                    new { IdOt = idOt, IdLaboratorio = idLaboratorio, Correlativo = correlativo },
                    commandType: CommandType.StoredProcedure // Indicar que es un SP
                );
            }
            
        }
        
        public IEnumerable<T> ViasResultados<T>(int idOt, int correlativo)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string storedProcedure = "usp_ListarViasInformeEnsayo"; // Nombre del procedimiento almacenado
        
                return connection.Query<T>(
                    storedProcedure,
                    new { IdOT = idOt, Correlativo = correlativo },
                    commandType: CommandType.StoredProcedure // Indicar que es un SP
                );
            }
        }
        
        public IEnumerable<T> ObtenerMuestras<T>(int idOt, int correlativo)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string storedProcedure = "usp_get_ObtenerViasDetalleInf"; // Nombre del procedimiento almacenado
        
                return connection.Query<T>(
                    storedProcedure,
                    new { IdOt = idOt, Correlativo = correlativo },
                    commandType: CommandType.StoredProcedure // Indicar que es un SP
                );
            }
        }
        
    }
}