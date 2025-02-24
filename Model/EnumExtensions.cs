using System;
using System.Reflection;

[AttributeUsage(AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
sealed class EnumDescriptionAttribute : Attribute
{
    public string Description { get; }

    public EnumDescriptionAttribute(string description)
    {
        Description = description;
    }
}

public static class EnumExtensions
{
    public static string ToDescription(this Enum value)
    {
        FieldInfo field = value.GetType().GetField(value.ToString());
        EnumDescriptionAttribute attribute = (EnumDescriptionAttribute)field.GetCustomAttribute(typeof(EnumDescriptionAttribute));
        return attribute == null ? value.ToString() : attribute.Description;
    }
}