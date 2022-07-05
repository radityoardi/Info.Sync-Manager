using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace System.DirectoryServices.AccountManagement
{
	[DirectoryRdnPrefix("CN")]
	[DirectoryObjectClass("user")]
	public class UserExtPrincipal : UserPrincipal
	{
		private ExtAdvancedFilters _extAdvancedFilters;
		public UserExtPrincipal(PrincipalContext context) : base(context) { }
		public UserExtPrincipal(PrincipalContext context, string samAccountName, string Password, bool Enabled) : base(context, samAccountName, Password, Enabled) { }
		public bool IsAttributeMulti(string attribute)
		{
			if (this.ExtensionGet(attribute).Length > 1)
			{
				return true;
			}
			return false;
		}
		public object AttributeGet(string attribute)
		{
			var x = this.ExtensionGet(attribute);
			if (x.Length == 1)
			{
				return this.ExtensionGet(attribute)[0];
			}
			return null;
		}
		public object[] AttributeGetMulti(string attribute)
		{
			return this.ExtensionGet(attribute);
		}
		public void AttributeSet(string attribute, object value)
		{
			this.ExtensionSet(attribute, value);
		}
		public ExtAdvancedFilters ExtAdvancedFilters
		{
			get
			{
				return this.AdvancedSearchFilter as ExtAdvancedFilters;
			}
		}
		public override AdvancedFilters AdvancedSearchFilter
		{
			get
			{
				if (_extAdvancedFilters == null)
				{
					_extAdvancedFilters = new ExtAdvancedFilters(this);
				}
				return _extAdvancedFilters;
			}
		}
	}
	public class ExtAdvancedFilters : AdvancedFilters
	{
		public ExtAdvancedFilters(Principal principal) : base(principal) { }
		public void IsNotPresent(string attribute)
		{
			this.IsNotPresent(attribute, typeof(string));
		}
		public void IsNotPresent(string attribute, Type type)
		{
			this.AdvancedFilterSet(attribute, "*", type, MatchType.NotEquals);//(!employeeID=*)
		}
		public void IsPresent(string attribute)
		{
			this.IsPresent(attribute, typeof(string));
		}
		public void IsPresent(string attribute, Type type)
		{
			this.AdvancedFilterSet(attribute, "*", type, MatchType.Equals);
		}
	}
}
