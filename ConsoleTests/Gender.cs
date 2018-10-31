using System.ComponentModel.DataAnnotations;

namespace ConsoleTests
{
    public enum Gender
    {
        [Display(Name = "Не указан")] Undefined,

        [Display(Name = "Мужской")] Male,

        [Display(Name = "Женский")] Female
    }
}