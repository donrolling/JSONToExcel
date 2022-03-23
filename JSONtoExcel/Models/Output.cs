using System.ComponentModel.DataAnnotations;

namespace JSONtoExcel.Models
{
	public class Output
	{
		[Required]
		[Display(Name = "First Name")]
		public string FirstName { get; set; }

		[Display(Name = "Middle Name")]
		public string MiddleName { get; set; }

		[Required]
		[Display(Name = "Last Name")]
		public string LastName { get; set; }

		[Display(Name = "Suffix")]
		public string Suffix { get; set; }

		[Display(Name = "Court Name")]
		public string CourtName { get; set; }

		[Required]
		[Display(Name = "City")]
		public string City { get; set; }

		[Required]
		[Display(Name = "State")]
		public string State { get; set; }

		[Required]
		[Display(Name = "Services")]
		public string Services { get; set; }

		[Required]
		[Display(Name = "Reviewer Name")]
		public string ReviewerName { get; set; }

		[Required]
		[Display(Name = "Reviewer Username")]
		public string ReviewerUsername { get; set; }

		[Required]
		[Display(Name = "Submitted Date")]
		public string SubmittedDate { get; set; }

		[Required]
		[Display(Name = "Littler Private")]
		public string LittlerPrivate { get; set; }

		[Required]
		[Display(Name = "Basic Review")]
		public string BasicReview { get; set; }

		[Required]
		[Display(Name = "Overall Performance")]
		public string OverallPerformance { get; set; }

		[Required]
		[Display(Name = "Average Rating")]
		public string AverageRating { get; set; }

		[Required]
		[Display(Name = "Skills Expertise Performance")]
		public string SkillsExpertisePerformance { get; set; }
	}
}