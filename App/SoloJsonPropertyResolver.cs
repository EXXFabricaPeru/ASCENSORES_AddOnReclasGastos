using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

public class SoloJsonPropertyResolver : DefaultContractResolver
{
    protected override IList<JsonProperty> CreateProperties(Type type, MemberSerialization memberSerialization)
    {
        // Filtra las propiedades para incluir solo aquellas con el decorador [JsonProperty]
        var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Where(p => p.GetCustomAttribute<JsonPropertyAttribute>() != null)
                        .Select(p => base.CreateProperty(p, memberSerialization))
                        .ToList();

        props.ForEach(p => { p.Writable = true; p.Readable = true; });
        return props;
    }
}