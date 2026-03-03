using System;
using System.ComponentModel;

namespace StockManager.Models
{
        public class StockInfo : INotifyPropertyChanged
        {
                private string _ticker;
                private string _name;
                private double? _price;
                private double? _changePercent;
                private double? _previousClose;
                private string _source;
                private DateTime? _updatedAt;

                public string Ticker
                {
                        get => _ticker;
                        set
                        {
                                _ticker = value;
                                OnPropertyChanged(nameof(Ticker));
                        }
                }

                public string Name
                {
                        get => _name;
                        set
                        {
                                _name = value;
                                OnPropertyChanged(nameof(Name));
                        }
                }

                public double? Price
                {
                        get => _price;
                        set
                        {
                                _price = value;
                                OnPropertyChanged(nameof(Price));
                        }
                }

                public double? ChangePercent
                {
                        get => _changePercent;
                        set
                        {
                                _changePercent = value;
                                OnPropertyChanged(nameof(ChangePercent));
                        }
                }

                public double? PreviousClose
                {
                        get => _previousClose;
                        set
                        {
                                _previousClose = value;
                                OnPropertyChanged(nameof(PreviousClose));
                        }
                }

                public string Source
                {
                        get => _source;
                        set
                        {
                                _source = value;
                                OnPropertyChanged(nameof(Source));
                        }
                }

                public DateTime? UpdatedAt
                {
                        get => _updatedAt;
                        set
                        {
                                _updatedAt = value;
                                OnPropertyChanged(nameof(UpdatedAt));
                        }
                }

                public StockInfo(string ticker, string name)
                {
                        Ticker = ticker;
                        Name = name;
                        Source = "等待更新";
                }

                public event PropertyChangedEventHandler PropertyChanged;

                protected virtual void OnPropertyChanged(string propertyName)
                {
                        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                }
        }
}
