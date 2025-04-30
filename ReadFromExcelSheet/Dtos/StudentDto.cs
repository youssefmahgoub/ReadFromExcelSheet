using System.ComponentModel.DataAnnotations;

public class StudentDto
{
    [Required]
    public string Name { get; set; }

    [Range(18, 60)]
    public int Age { get; set; }

    [EmailAddress]
    public string Email { get; set; }
    [Required]
    [MinLength(1)]
    public byte[] ProfilePicture { get; set; }
}