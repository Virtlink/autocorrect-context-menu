using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCorrectContextMenu
{
	/// <summary>
	/// A spelling correction.
	/// </summary>
	public struct SpellingCorrection
	{
		/// <summary>
		/// Gets the word to be corrected.
		/// </summary>
		public string Word { get; }

		/// <summary>
		/// Gets the suggested correction.
		/// </summary>
		public string Correction { get; }

		#region Constructors
		/// <summary>
		/// Initializes a new instance of the <see cref="SpellingCorrection"/> class.
		/// </summary>
		public SpellingCorrection(string word, string correction)
		{
			this.Word = word.Trim();
			this.Correction = correction.Trim();
		}
		#endregion

		/// <summary>
		/// Encodes the correction as a string.
		/// </summary>
		/// <param name="correction">The correction to encode.</param>
		/// <returns>The encoded correction.</returns>
		public static string Encode(SpellingCorrection correction)
		{
			return $"{Base64Encode(correction.Word)}|{Base64Encode(correction.Correction)}";
		}

		/// <summary>
		/// Decodes the correction from a string.
		/// </summary>
		/// <param name="encoded">The correction to decode.</param>
		/// <returns>The decoded correction.</returns>
		public static SpellingCorrection Decode(string encoded)
		{
			int split = encoded.IndexOf("|", StringComparison.Ordinal);
			string word = Base64Decode(encoded.Substring(0, split));
			string correction = Base64Decode(encoded.Substring(split + 1));
			return new SpellingCorrection(word, correction);
		}

		private static string Base64Encode(string plainText)
		{
			var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
			return System.Convert.ToBase64String(plainTextBytes);
		}

		private static string Base64Decode(string base64EncodedData)
		{
			var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
			return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
		}

		/// <inheritdoc />
		public override string ToString()
		{
			return $"{this.Word} -> {this.Correction}";
		}
	}
}
