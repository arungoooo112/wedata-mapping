import * as vscode from 'vscode';
import * as yaml from 'js-yaml';
import * as fs from 'fs';
import * as path from 'path';
import * as exceljs from 'exceljs';

interface Mapping {
  folder: string;
  update: string;
  with: {
    set: string;
    case: {
      when: string;
      then: string;
    }[];
    where: {
      column: string;
      equals: string;
    };
  }[];
}

export function activate(context: vscode.ExtensionContext) {
  let disposable = vscode.commands.registerCommand('extension.executeMapping', async () => {
    try {
      // Prompt the user to select the mapping YAML file
      const mappingFilePath = await vscode.window.showOpenDialog({
        canSelectFiles: true,
        canSelectFolders: false,
        canSelectMany: false,
        filters: {
          YAML: ['yaml']
        }
      });

      if (!mappingFilePath || mappingFilePath.length === 0) {
        vscode.window.showErrorMessage('No mapping YAML file selected.');
        return;
      }

      // Load the mapping YAML file
      const yamlContent = fs.readFileSync(mappingFilePath[0].fsPath, 'utf8');
      const mapping: Mapping[] = yaml.load(yamlContent) as Mapping[];

      // Perform the mapping actions
      mapping.forEach(async (mappingItem) => {
        const excelFilePath = path.join(mappingItem.folder, 'your_excel_file.xlsx');

        const workbook = new exceljs.Workbook();
        await workbook.xlsx.readFile(excelFilePath);

        const worksheet = workbook.getWorksheet(mappingItem.update);

        if (worksheet) {
			mappingItem.with.forEach((withItem) => {
				worksheet.eachRow((row, rowNumber) => {
				  const columnValue = row.getCell(withItem.where.column).value;
				  if (columnValue === withItem.where.equals) {
					withItem.case.forEach((caseItem) => {
					  const updatedValue = withItem.set.replace(caseItem.when, caseItem.then);
					  row.getCell(withItem.where.column).value = updatedValue;
					});
				  }
				});
			  });

          await workbook.xlsx.writeFile(excelFilePath);
        }
      });

      vscode.window.showInformationMessage('Mapping executed successfully.');
    } catch (error) {
      vscode.window.showErrorMessage(`Error executing mapping: ${error}`);
    }
  });

  context.subscriptions.push(disposable);
}

export function deactivate() {}