using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LumenWorks.Framework.IO.Csv;
using OfficeOpenXml;
using SharpKml.Base;
using SharpKml.Dom;

namespace CsvToKml
{
    public class Location
    {
        public Vector Coordinates { get; set; }
        public string Name { get; set; }
        public string ParentName { get; set; }
        public Location Parent { get; set; }
        public Folder PlacemarkFolder { get; set; }
        public Folder PathFolder { get; set; }
        public StyleMapCollection StyleMap { get; set; }
        public StyleMapCollection StyleMapNoLabels { get; set; }
    }

    public class Link
    {
        public string Name { get; set; }
        public Location Location1 { get; set; }
        public Location Location2 { get; set; }
    }
    
    public enum LabelMode
    {
        None,
        All,
        Some
    }

    public class Map
    {
        public LabelMode Labels { get; set; } = LabelMode.All;
        public Dictionary<string, Location> Locations { get; set; } = new Dictionary<string, Location>();
        public List<Link> Links { get; set; } = new List<Link>();
        private List<Color> _defaultColors = new List<Color>
        {
            Color.Black,
            Color.Aqua,
            Color.LawnGreen,
            Color.Violet,
            Color.White,
            Color.Orange,
            Color.LightPink,
            Color.LightGreen,
            Color.Yellow,
            Color.LightSkyBlue,
            Color.Red
            
        };
        private int _colourIndex = 0;

        public Map (string locationFile, string linkFile)
        {
            LoadLocationsFromCsv(locationFile);
            LoadLinksFromCsv(linkFile);
        }

        public Map(string file)
        {
            if (file.EndsWith("xls") || file.EndsWith("xlsm") || file.EndsWith("xlsx"))
            {
                LoadFromXls(file);
            }
            else
            {
                LoadLocationsFromCsv(file);
            }
        }

        private void LoadLocationsFromCsv(string file)
        {
            using (CachedCsvReader csv = new CachedCsvReader(new StreamReader(file), true))
            {
                while (csv.ReadNextRecord())
                {
                    float fLat = float.Parse(csv["Lat"]);
                    float fLong = float.Parse(csv["Long"]);
                    string name = csv["Name"];
                    string parent = csv["Group"];

                    Location l = new Location()
                    {
                        Name = name,
                        ParentName = parent,
                        Coordinates = new Vector(fLat, fLong)
                    };
                    Locations.Add(l.Name, l);
                }
            }
        }

        private void LoadLinksFromCsv(string file)
        {
            using (CachedCsvReader csv = new CachedCsvReader(new StreamReader(file), true))
            {
                while (csv.ReadNextRecord())
                {
                    try
                    {
                        Location l1 = Locations[csv["Location1"]];
                        Location l2 = Locations[csv["Location2"]];
                        string name = csv["Name"];

                        Link link = new Link()
                        {
                            Name = name,
                            Location1 = l1,
                            Location2 = l2
                        };
                        Links.Add(link);
                    }
                    catch { }
                }
            }
        }

        private void LoadFromXls(string file)
        {
            ExcelPackage pck = new ExcelPackage(new FileInfo(file));
            var sites = pck.Workbook.Worksheets["Locations"];
            var links = pck.Workbook.Worksheets["Links"];
            
            int nameIndex = 0;
            int latIndex = 0;
            int longIndex = 0;
            int repIndex = 0;

            for (int i = 1; i < 100; i++)
            {
                if (sites.Cells[1, i].Value.ToString() == "Name")
                    nameIndex = i;
                if (sites.Cells[1, i].Value.ToString() == "Lat")
                    latIndex = i;
                if (sites.Cells[1, i].Value.ToString() == "Long")
                    longIndex = i;
                if (sites.Cells[1, i].Value.ToString() == "Group")
                    repIndex = i;
                if (nameIndex != 0 && latIndex != 0 && longIndex != 0 && repIndex != 0)
                    break;
            }
            if (nameIndex == 0 || latIndex == 0 || longIndex == 0 || repIndex == 0)
                throw new Exception("Could not locate data");
            for (int i = 2; i <= sites.Dimension.End.Row; i++)
            {
                try
                {
                    float fLat = float.Parse(sites.Cells[i,latIndex].Value.ToString());
                    float fLong = float.Parse(sites.Cells[i, longIndex].Value.ToString());
                    string name = sites.Cells[i, nameIndex].Value.ToString();
                    string parent = sites.Cells[i, repIndex].Value?.ToString() ?? "";

                    Location l = new Location()
                    {
                        Name = name,
                        ParentName = parent,
                        Coordinates = new Vector(fLat, fLong)
                    };
                    Locations.Add(l.Name, l);
                }
                catch { }
            }
            if (links != null)
            {
                int lnameIndex = 0;
                int location1 = 0;
                int location2 = 0;

                for (int i = 1; i < 100; i++)
                {
                    if (links.Cells[1, i].Value.ToString() == "Name")
                        lnameIndex = i;
                    if (links.Cells[1, i].Value.ToString() == "Location1")
                        location1 = i;
                    if (links.Cells[1, i].Value.ToString() == "Location2")
                        location2 = i;
                    if (lnameIndex != 0 && location1 != 0 && location2 != 0)
                        break;
                }
                if (lnameIndex == 0 || location1 == 0 || location2 == 0)
                    throw new Exception("Could not locate data");
                for (int i = 2; i <= links.Dimension.End.Row; i++)
                {
                    try
                    {
                        Location l1 = Locations[links.Cells[i, location1].Value.ToString()];
                        Location l2 = Locations[links.Cells[i, location2].Value.ToString()];
                        string name = links.Cells[i, lnameIndex].Value.ToString();

                        Link l = new Link()
                        {
                            Name = name,
                            Location1 = l1,
                            Location2 = l2
                        };
                        Links.Add(l);
                    }
                    catch { }
                }
            }
                         
        }

        public void ProduceKml(string file)
        {
            Kml doc = new Kml();
            Document d = new Document() { Name = Path.GetFileName(file) };
            doc.Feature = d;
            
            Folder root = new Folder() { Name = "Network" };
            d.AddFeature(root);
            StyleMapCollection style_default = GenerateNewStyle("default", d, null);

            List<string> folders = (from i in Locations.Values where !string.IsNullOrWhiteSpace(i.ParentName) select i.ParentName ).Distinct().ToList();
            foreach (var str in folders)
            {
                Folder fPlacemark = new Folder() { Name = str };
                Folder fPath = new Folder() { Name = "Paths" };
                root.AddFeature(fPlacemark);
                if (!Locations.ContainsKey(str))
                {
                    Locations.Add(str, new Location() { Name = str, PlacemarkFolder=fPlacemark });
                }
                else
                {
                    fPlacemark.AddFeature(fPath);
                    Locations[str].PathFolder = fPath;
                    Locations[str].PlacemarkFolder = fPlacemark;
                }
                GenerateNewStyle(str.Sanitize(), d, Locations[str]);
            }
            foreach (var loc in Locations.Values)
            {
                var parent = (from i in Locations.Values where i.Name == loc.ParentName select i).FirstOrDefault();
                loc.Parent = parent;
            }
            
            foreach (var loc in Locations.Values)
            {
                if (loc.Coordinates == null)
                    continue;
                Placemark plPoint = new Placemark();
                Placemark plPath = new Placemark();

                if (loc.Parent != null)
                {
                    switch (Labels)
                    {
                        case LabelMode.All:
                            plPoint.StyleUrl = new Uri($"#msn_{loc.ParentName.Sanitize()}", UriKind.Relative);
                            break;
                        case LabelMode.Some:
                            if (IsImportant(loc.Name))
                            {
                                plPoint.StyleUrl = new Uri($"#msn_{loc.ParentName.Sanitize()}", UriKind.Relative);
                            }
                            else
                            {
                                plPoint.StyleUrl = new Uri($"#msn_{loc.ParentName.Sanitize()}_nolabels", UriKind.Relative);
                            }
                            break;
                        case LabelMode.None:
                            plPoint.StyleUrl = new Uri($"#msn_{loc.ParentName.Sanitize()}_nolabels", UriKind.Relative);
                            break;
                    }
                    
                }
                else if (loc.StyleMap != null)
                {
                    switch (Labels)
                    {
                        case LabelMode.All:
                        case LabelMode.Some:
                            plPoint.StyleUrl = new Uri($"#{loc.StyleMap.Id}", UriKind.Relative);
                            break;
                        case LabelMode.None:
                            plPoint.StyleUrl = new Uri($"#{loc.StyleMapNoLabels.Id}", UriKind.Relative);
                            break;
                    }
                }
                else
                {
                    plPoint.StyleUrl = new Uri($"#{style_default.Id}", UriKind.Relative);
                }
                plPoint.Name = loc.Name;
                plPoint.Geometry = new SharpKml.Dom.Point() { Coordinate = loc.Coordinates };
                Folder placemarkFolder = loc.Parent?.PlacemarkFolder ?? loc.PlacemarkFolder ?? root;
                placemarkFolder.AddFeature(plPoint);
                if (loc.Parent?.Coordinates != null)
                {
                    plPath.StyleUrl = new Uri($"#msn_{loc.ParentName.Sanitize()}", UriKind.Relative);
                    plPath.Name = loc.Name + " Path";
                    LineString l = new LineString();
                    l.Coordinates = new CoordinateCollection();
                    l.Coordinates.Add(loc.Coordinates);
                    l.Coordinates.Add(loc.Parent.Coordinates);
                    plPath.Geometry = l;
                    loc.Parent.PathFolder.AddFeature(plPath);
                }                
            }
            Folder linksFolder = new Folder() { Name = "Links" };
            Style s = new Style()
            {
                Id = $"sn_links",
                Line = new LineStyle()
                {
                    Color = new Color32(255, 255, 0, 0),
                    Width = 4.0
                }
            };
            root.AddStyle(s);
            if (Links.Count > 0)
            {
                root.AddFeature(linksFolder);
            }
            foreach (Link link in Links)
            {
                Placemark p = new Placemark() { Name = link.Name };
                LineString l = new LineString();
                l.Coordinates = new CoordinateCollection();
                l.Coordinates.Add(link.Location1.Coordinates);
                l.Coordinates.Add(link.Location2.Coordinates);
                p.Geometry = l;
                linksFolder.AddFeature(p);
                p.StyleUrl = new Uri($"#{s.Id}", UriKind.Relative);
            }
            Serializer serializer = new Serializer();
            serializer.Serialize(doc);
            File.WriteAllText(file, serializer.Xml);
        }

        private bool IsImportant(string name)
        {
            if (name.Contains(" ZS")) return true;
            if (name.Contains(" SS")) return true;
            if (name.Contains(" Gen")) return true;
            if (name.Contains(" GXP")) return true;
            return false;
        }

        private StyleMapCollection GenerateNewStyle(string str, Feature feature, Location loc)
        {
            Color32 c = GetRandomColor();
            StyleMapCollection stylemap = new StyleMapCollection();
            StyleMapCollection stylemap2 = new StyleMapCollection();
            Style s_normal = new Style()
            {
                Icon = new IconStyle
                {
                    Icon = new IconStyle.IconLink(new Uri("http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png")),
                    Color = c
                },
                Id = $"sn_{str}",
                Label = new LabelStyle()
                {
                    Scale = 0.5
                },
                Line = new LineStyle()
                {
                    Color = new Color32(255 / 2, c.Blue, c.Green, c.Red)
                }
            };
            Style s_nolabels = new Style()
            {
                Icon = new IconStyle
                {
                    Icon = new IconStyle.IconLink(new Uri("http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png")),
                    Color = c
                },
                Id = $"sn_{str}_nolabels",
                Label = new LabelStyle()
                {
                    Scale = 0.01
                },
                Line = new LineStyle()
                {
                    Color = new Color32(255 / 2, c.Blue, c.Green, c.Red)
                }
            };
            Style s_highlight = new Style()
            {
                Icon = new IconStyle
                {
                    Icon = new IconStyle.IconLink(new Uri("http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png")),
                    Color = c
                },
                Line = new LineStyle()
                {
                    Color = new Color32(255, c.Blue, c.Green, c.Red),
                    Width = 4
                },
                Id = $"sn_{str}_highlight",
            };
            stylemap.Id = $"msn_{str}"; 
            stylemap.Add(new Pair() { State = StyleState.Normal, StyleUrl = new Uri($"#{s_normal.Id}", UriKind.Relative)});
            stylemap.Add(new Pair() { State = StyleState.Highlight, StyleUrl = new Uri($"#{s_highlight.Id}", UriKind.Relative)});
            stylemap2.Id = $"msn_{str}_nolabels";
            stylemap2.Add(new Pair() { State = StyleState.Normal, StyleUrl = new Uri($"#{s_nolabels.Id}", UriKind.Relative) });
            stylemap2.Add(new Pair() { State = StyleState.Highlight, StyleUrl = new Uri($"#{s_highlight.Id}", UriKind.Relative) });
            feature.AddStyle(stylemap);
            feature.AddStyle(stylemap2);
            feature.AddStyle(s_normal);
            feature.AddStyle(s_nolabels);
            feature.AddStyle(s_highlight);
            if (loc != null)
            {
                loc.StyleMap = stylemap;
                loc.StyleMapNoLabels = stylemap2;
            }
            return Labels == LabelMode.All ? stylemap : stylemap2;
        }

        private Color32 GetRandomColor()
        {
            Color c = _defaultColors[_colourIndex];
            _colourIndex = ++_colourIndex % _defaultColors.Count;
            return new Color32(255, c.B, c.G, c.R);
        }
    }

}
