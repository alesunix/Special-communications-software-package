using System;
using System.ComponentModel.DataAnnotations;
using System.Data.Linq.Mapping;


namespace ProgramCCS
{
    [Table(Name = "Table_1")]
    public class Orders_incomplete
    {
        [Column(Name = "id", IsPrimaryKey = true, IsDbGenerated = true)]
        public int ID { get; set; }
        [Column(Name = "oblast")]
        public string Область { get; set; }
        [Column(Name = "punkt")]
        public string Населенный_пункт { get; set; }
        [Column(Name = "familia")]
        public string Ф_И_О { get; set; }
        [Column(Name = "summ")]
        public double Стоимость { get; set; }
        [Column(Name = "plata_za_uslugu")]
        public double Услуга { get; set; }
        [Column(Name = "tarif")]
        public int Тариф { get; set; }
        [Column(Name = "doplata")]
        public int Доплата { get; set; }
        [Column(Name = "ob_cennost")]
        public int Оц { get; set; }
        [Column(Name = "plata_za_nalog")]
        public double Нп { get; set; }
        [Column(Name = "N_zakaza")]
        public string N_Заказа { get; set; }
        [Column(Name = "status")]
        public string Статус { get; set; }
        [Column(Name = "data_zapisi")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{dd/MM/yyyy}")]
        public DateTime Дата_записи { get; set; }
        [Column(Name = "prichina")]
        public string Причина { get; set; }
        [Column(Name = "obrabotka")]
        public string Обработка { get; set; }
        [Column(Name = "data_obrabotki")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{dd/MM/yyyy}")]
        public DateTime? Дата_обработки { get; set; }
        [Column(Name = "filial")]
        public string Филиал { get; set; }
        [Column(Name = "client")]
        public string Контрагент { get; set; }
        [Column(Name = "nomer_spiska")]
        public string Список { get; set; }
        [Column(Name = "nomer_nakladnoy")]
        public string Накладная { get; set; }
        [Column(Name = "nomer_reestra")]
        public string Реестр { get; set; }
        //[Column(Name = "Ns")]
        //public int Ns { get; set; }
        //[Column(Name = "Nn")]
        //public int Nn { get; set; }
        //[Column(Name = "Nr")]
        //public int Nr { get; set; }
        [Column(Name = "tarifs")]
        public string Тарифы { get; set; }
    }

   
}
