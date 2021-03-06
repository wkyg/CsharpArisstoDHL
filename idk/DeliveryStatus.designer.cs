#pragma warning disable 1591
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace idk
{
	using System.Data.Linq;
	using System.Data.Linq.Mapping;
	using System.Data;
	using System.Collections.Generic;
	using System.Reflection;
	using System.Linq;
	using System.Linq.Expressions;
	using System.ComponentModel;
	using System;
	
	
	[global::System.Data.Linq.Mapping.DatabaseAttribute(Name="NEP")]
	public partial class DeliveryStatusDataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region Extensibility Method Definitions
    partial void OnCreated();
    partial void InsertDelivery_Status_WK(Delivery_Status_WK instance);
    partial void UpdateDelivery_Status_WK(Delivery_Status_WK instance);
    partial void DeleteDelivery_Status_WK(Delivery_Status_WK instance);
    #endregion
		
		public DeliveryStatusDataContext() : 
				base(global::idk.Properties.Settings.Default.NEPConnectionString, mappingSource)
		{
			OnCreated();
		}
		
		public DeliveryStatusDataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DeliveryStatusDataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DeliveryStatusDataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DeliveryStatusDataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<Delivery_Status_WK> Delivery_Status_WKs
		{
			get
			{
				return this.GetTable<Delivery_Status_WK>();
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.Delivery_Status_WK")]
	public partial class Delivery_Status_WK : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _ID;
		
		private string _SinNo;
		
		private string _DeliveryComp;
		
		private string _TrackingNo;
		
		private string _DeliveryStatus;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnIDChanging(int value);
    partial void OnIDChanged();
    partial void OnSinNoChanging(string value);
    partial void OnSinNoChanged();
    partial void OnDeliveryCompChanging(string value);
    partial void OnDeliveryCompChanged();
    partial void OnTrackingNoChanging(string value);
    partial void OnTrackingNoChanged();
    partial void OnDeliveryStatusChanging(string value);
    partial void OnDeliveryStatusChanged();
    #endregion
		
		public Delivery_Status_WK()
		{
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ID", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int ID
		{
			get
			{
				return this._ID;
			}
			set
			{
				if ((this._ID != value))
				{
					this.OnIDChanging(value);
					this.SendPropertyChanging();
					this._ID = value;
					this.SendPropertyChanged("ID");
					this.OnIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_SinNo", DbType="VarChar(MAX)")]
		public string SinNo
		{
			get
			{
				return this._SinNo;
			}
			set
			{
				if ((this._SinNo != value))
				{
					this.OnSinNoChanging(value);
					this.SendPropertyChanging();
					this._SinNo = value;
					this.SendPropertyChanged("SinNo");
					this.OnSinNoChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DeliveryComp", DbType="VarChar(MAX)")]
		public string DeliveryComp
		{
			get
			{
				return this._DeliveryComp;
			}
			set
			{
				if ((this._DeliveryComp != value))
				{
					this.OnDeliveryCompChanging(value);
					this.SendPropertyChanging();
					this._DeliveryComp = value;
					this.SendPropertyChanged("DeliveryComp");
					this.OnDeliveryCompChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_TrackingNo", DbType="VarChar(MAX) NOT NULL", CanBeNull=false)]
		public string TrackingNo
		{
			get
			{
				return this._TrackingNo;
			}
			set
			{
				if ((this._TrackingNo != value))
				{
					this.OnTrackingNoChanging(value);
					this.SendPropertyChanging();
					this._TrackingNo = value;
					this.SendPropertyChanged("TrackingNo");
					this.OnTrackingNoChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DeliveryStatus", DbType="VarChar(MAX)")]
		public string DeliveryStatus
		{
			get
			{
				return this._DeliveryStatus;
			}
			set
			{
				if ((this._DeliveryStatus != value))
				{
					this.OnDeliveryStatusChanging(value);
					this.SendPropertyChanging();
					this._DeliveryStatus = value;
					this.SendPropertyChanged("DeliveryStatus");
					this.OnDeliveryStatusChanged();
				}
			}
		}
		
		public event PropertyChangingEventHandler PropertyChanging;
		
		public event PropertyChangedEventHandler PropertyChanged;
		
		protected virtual void SendPropertyChanging()
		{
			if ((this.PropertyChanging != null))
			{
				this.PropertyChanging(this, emptyChangingEventArgs);
			}
		}
		
		protected virtual void SendPropertyChanged(String propertyName)
		{
			if ((this.PropertyChanged != null))
			{
				this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
			}
		}
	}
}
#pragma warning restore 1591
