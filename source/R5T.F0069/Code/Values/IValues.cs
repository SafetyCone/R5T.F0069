using System;

using R5T.T0131;


namespace R5T.F0069
{
	[ValuesMarker]
	public partial interface IValues : IValuesMarker
	{
		public bool DefaultApplicationVisibility => true;
	}
}