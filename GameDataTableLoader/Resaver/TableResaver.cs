using GameDataTableLoader.Loader;

namespace GameDataTableLoader.Resaver
{
	public class TableResaver<T>
		where T : class, new()
	{
		private TableLoader<T> _tableLoader;

		public TableResaver(TableLoader<T> tableLoader)
		{
			_tableLoader = tableLoader;
		}

		public void Run()
		{
			_tableLoader.Run();
			var data = _tableLoader.TableInfos();
			foreach (var info in data)
			{
				info.Value.Save();
			}
		}
	}
}
