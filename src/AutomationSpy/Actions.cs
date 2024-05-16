
using System.Windows;
using System.Windows.Controls;

namespace dDeltaSolutions.Spy
{
    public partial class MainWindow
    {
		private void OnActions(object sender, RoutedEventArgs e)
		{
			TreeViewItem treeviewItem = this.tvElements.SelectedItem as TreeViewItem;
            if (treeviewItem == null)
            {
				System.Windows.MessageBox.Show("Select an element in the tree");
                return;
            }
			
			TreeNode node = treeviewItem.Tag as TreeNode;
            if (node == null)
            {
                return;
            }
			
			if (node.IsAlive == false)
			{
				System.Windows.MessageBox.Show("The selected element is not available anymore");
				return;
			}
			
			bool timerStopped = false;
			if (timer != null && timer.Enabled == true)
            {
                timer.Enabled = false;
				timerStopped = true;
            }
			
			bool timerTrackStopped = false;
			if (timerTrack != null && timerTrack.Enabled == true)
            {
                timerTrack.Enabled = false;
				timerTrackStopped = true;
            }
			
			WindowActions wndActions = new WindowActions(node);
            wndActions.Owner = this;
			wndActions.ShowDialog();
			
			if (timerStopped == true)
			{
				timer.Enabled = true;
			}
			
			if (timerTrackStopped == true)
			{
				timerTrack.Enabled = true;
			}
		}
	}
}