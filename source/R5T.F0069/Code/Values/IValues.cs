using System;

using R5T.T0131;


namespace R5T.F0069
{
	[ValuesMarker]
	public partial interface IValues : IValuesMarker
	{
		/// <summary>
		/// <para><value>true</value></para>
		/// </summary>
		public bool ApplicationVisibility_Default => true;
	}
}