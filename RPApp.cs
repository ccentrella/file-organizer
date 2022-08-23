using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Academic_Pro
{
	public class RPApp
	{
		public RPApp(string location, string name)
		{
			AppLocation = location;
			Title= name;
		}

		public string AppLocation{ get; private set; }

		public string  Title { get; private set; }
	}
}
