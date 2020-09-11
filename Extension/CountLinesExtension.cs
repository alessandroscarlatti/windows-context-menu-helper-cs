using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
/// <summary>
/// The CountLinesExtensions is an example shell context menu extension,
/// implemented with SharpShell. It adds the command 'Count Lines' to text
/// files.
/// </summary>
[ComVisible(true)]
[COMServerAssociation(AssociationType.Class, "*")]
[COMServerAssociation(AssociationType.Class, "Directory\\Background")]
[DisplayName("Count Lines Handler")]
public class CountLinesExtension : SharpContextMenu
{

    private ToolStripMenuItem itemCountLines;


    /// <summary>
    /// Determines whether this instance can a shell
    /// context show menu, given the specified selected file list
    /// </summary>
    /// <returns>
    /// <c>true</c> if this instance should show a shell context
    /// menu for the specified file list; otherwise, <c>false</c>
    /// </returns>
    protected override bool CanShowMenu()
    {
        //  We always show the menu
        return true;
    }

    /// <summary>
    /// Creates the context menu. This can be a single menu item or a tree of them.
    /// </summary>
    /// <returns>
    /// The context menu for the shell context menu.
    /// </returns>
    protected override ContextMenuStrip CreateMenu()
    {
        return _CreateNestedContextMenu();
    }

    private ContextMenuStrip _CreateNestedContextMenu()
    {
        var rootMenu = new ContextMenuStrip();

        var menu = new ToolStripMenuItem("Menu 1");

        var childMenu1 = new ToolStripMenuItem("Child Menu 1");
        var childMenu2 = new ToolStripMenuItem("Child Menu 2");

        var childCommand1_1 = new ToolStripMenuItem("Child Command 1.1");
        var childCommand1_2 = new ToolStripMenuItem("Child Command 1.2");
        var childCommand1_3 = new ToolStripMenuItem("Child Command 1.3");
        var childCommand1_4 = new ToolStripMenuItem("Child Command 1.4");
        var childCommand1_5 = new ToolStripMenuItem("Child Command 1.5");
        var childCommand1_6 = new ToolStripMenuItem("Child Command 1.6");
        var childCommand1_7 = new ToolStripMenuItem("Child Command 1.7");
        var childCommand1_8 = new ToolStripMenuItem("Child Command 1.8");
        var childCommand1_9 = new ToolStripMenuItem("Child Command 1.9");

        var childCommand2_1 = new ToolStripMenuItem("Child Command 2.1");
        var childCommand2_2 = new ToolStripMenuItem("Child Command 2.2");
        var childCommand2_3 = new ToolStripMenuItem("Child Command 2.3");
        var childCommand2_4 = new ToolStripMenuItem("Child Command 2.4");
        var childCommand2_5 = new ToolStripMenuItem("Child Command 2.5");
        var childCommand2_6 = new ToolStripMenuItem("Child Command 2.6");
        var childCommand2_7 = new ToolStripMenuItem("Child Command 2.7");
        var childCommand2_8 = new ToolStripMenuItem("Child Command 2.8");
        var childCommand2_9 = new ToolStripMenuItem("Child Command 2.9");

        childMenu1.DropDownItems.Add(childCommand1_1);
        childMenu1.DropDownItems.Add(childCommand1_2);
        childMenu1.DropDownItems.Add(childCommand1_3);
        childMenu1.DropDownItems.Add(childCommand1_4);
        childMenu1.DropDownItems.Add(childCommand1_5);
        childMenu1.DropDownItems.Add(childCommand1_6);
        childMenu1.DropDownItems.Add(childCommand1_7);
        childMenu1.DropDownItems.Add(childCommand1_8);
        childMenu1.DropDownItems.Add(childCommand1_9);

        childMenu2.DropDownItems.Add(childCommand2_1);
        childMenu2.DropDownItems.Add(childCommand2_2);
        childMenu2.DropDownItems.Add(childCommand2_3);
        childMenu2.DropDownItems.Add(childCommand2_4);
        childMenu2.DropDownItems.Add(childCommand2_5);
        childMenu2.DropDownItems.Add(childCommand2_6);
        childMenu2.DropDownItems.Add(childCommand2_7);
        childMenu2.DropDownItems.Add(childCommand2_8);
        childMenu2.DropDownItems.Add(childCommand2_9);

        menu.DropDownItems.Add(childMenu1);
        menu.DropDownItems.Add(childMenu2);

        rootMenu.Items.Add(menu);

        var menu2 = new ToolStripMenuItem("Command 2");
        rootMenu.Items.Add(menu2);

        var menu3 = new ToolStripMenuItem("Menu 3");

        var childMenu3 = new ToolStripMenuItem("Child Menu 3");
        var childCommand3_1 = new ToolStripMenuItem("Child Command 3.1");
        var childCommand3_2 = new ToolStripMenuItem("Child Command 3.2");
        var childCommand3_3 = new ToolStripMenuItem("Child Command 3.3");
        childMenu3.DropDownItems.Add(childCommand3_1);
        childMenu3.DropDownItems.Add(childCommand3_2);
        childMenu3.DropDownItems.Add(childCommand3_3);

        menu3.DropDownItems.Add(childMenu3);
        rootMenu.Items.Add(menu3);

        return rootMenu;
    }

    private ContextMenuStrip _CreateCountLinesContextMenu()
    {
        //  Create the menu strip
        var menu = new ContextMenuStrip();

        //  Create a 'count lines' item
        itemCountLines = new ToolStripMenuItem
        {
            Text = string.Format("Count Lines for {0}...", GetName())
            // Image = Properties.Resources.CountLines
        };

        //  When we click, we'll count the lines
        itemCountLines.Click += (sender, args) => CountLines();

        //  Add the item to the context menu.
        menu.Items.Add(itemCountLines);

        //  Return the menu
        return menu;
    }

    protected override void OnInitialiseMenu(int parentItemIndex)
    {
        itemCountLines.Text = string.Format("Count Lines for init {0}...", GetName());
    }

    private string GetName()
    {
        foreach (var filePath in SelectedItemPaths)
        {
            //  get the first file name, if any selected
            return Path.GetFileName(filePath);
        }
        return "files";
    }

    /// <summary>
    /// Counts the lines in the selected files
    /// </summary>
    private void CountLines()
    {
        //  Builder for the output
        var builder = new StringBuilder();

        //  Go through each file.
        foreach (var filePath in SelectedItemPaths)
        {
            //  Count the lines
            builder.AppendLine(string.Format("{0} - {1} Lines",
              Path.GetFileName(filePath), File.ReadAllLines(filePath).Length));
        }

        //  Show the output
        MessageBox.Show(builder.ToString());
    }
}