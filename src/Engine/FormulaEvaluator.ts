import Cell from "./Cell";
import SheetMemory from "./SheetMemory";
import { ErrorMessages } from "./GlobalDefinitions";

type FormulaType = string[];
type TokenType = string;

export class FormulaEvaluator {
    private _errorOccurred: boolean = false;
    private _errorMessage: string = "";
    private _currentFormula: FormulaType = [];
    private _lastResult: number = 0;
    private _sheetMemory: SheetMemory;
    private _result: number = 0;

    constructor(memory: SheetMemory) {
        this._sheetMemory = memory;
    }

    evaluate(formulaInput: FormulaType) {
        this._errorMessage = "";
        this._result = 0; 
    
        if(formulaInput.length === 0) {
            this._errorMessage = ErrorMessages.emptyFormula;
            return;
        }
    
        const lastElementInFormulaInput = formulaInput[formulaInput.length - 1];
        if (['+', '-', '*', '/'].includes(lastElementInFormulaInput)) {
            this._errorMessage = ErrorMessages.invalidFormula;
            formulaInput = formulaInput.slice(0, -1); // Remove the last operator from the formula
        }
    
        try {
            let currentTokenIndex = 0;
            const getNextTokenInFormula = () => formulaInput[currentTokenIndex];
            const readAndMoveToNextToken = () => formulaInput[currentTokenIndex++];
    
            const extractNumberFromFormula = () => {
                if (this.isCellReference(getNextTokenInFormula())) {
                    const [extractedValue, errorFound] = this.getCellValue(readAndMoveToNextToken());
                    if (errorFound) {
                        this._errorMessage = errorFound;
                        throw new Error(errorFound);
                    }
                    return extractedValue;
                }
    
                if (!this.isNumber(getNextTokenInFormula())) {
                    this._errorMessage = ErrorMessages.invalidNumber;
                    throw new Error(ErrorMessages.invalidNumber);
                }
    
                this._lastResult = parseFloat(readAndMoveToNextToken());
                return this._lastResult;
            };
    
            const calculateTermValue = ():number => {
                let leftValue = calculateFactorValue();
    
                while (getNextTokenInFormula() === '*' || getNextTokenInFormula() === '/') {
                    const operatorFound = readAndMoveToNextToken();
                    const rightValue = calculateFactorValue();
    
                    if (operatorFound === '*') leftValue *= rightValue;
                    else {
                        if (rightValue === 0) {
                            this._errorMessage = ErrorMessages.divideByZero;
                            throw new Error(ErrorMessages.divideByZero);
                        }
                        leftValue /= rightValue;
                    }
                }
    
                return leftValue;
            };
    
            const evaluateExpressionValue = () => {
                let accumulatedValue = calculateTermValue();
    
                while (getNextTokenInFormula() === '+' || getNextTokenInFormula() === '-') {
                    const operatorFound = readAndMoveToNextToken();
                    const subsequentValue = calculateTermValue();
    
                    if (operatorFound === '+') accumulatedValue += subsequentValue;
                    else accumulatedValue -= subsequentValue;
                }
    
                return accumulatedValue;
            };
    
            const calculateFactorValue = () => {
                if (getNextTokenInFormula() === '(') {
                    readAndMoveToNextToken(); // Consume '('
                    const innerExpressionValue = evaluateExpressionValue();
    
                    if (readAndMoveToNextToken() !== ')') {
                        this._errorMessage = ErrorMessages.missingParentheses;
                        throw new Error(ErrorMessages.missingParentheses);
                    }
    
                    return innerExpressionValue;
                }
    
                return extractNumberFromFormula();
            };
    
            this._result = evaluateExpressionValue();
    
            if (currentTokenIndex < formulaInput.length) {
                this._errorMessage = ErrorMessages.invalidFormula;
                throw new Error(ErrorMessages.invalidFormula);
            }
        } catch (error) {
            this._errorOccurred = true;
    
            if(this._errorMessage === ErrorMessages.divideByZero) {
                this._result = Infinity;
            }
    
            if (this._errorOccurred && this._result === 0) {
                this._result = this._lastResult;
            }
        }
    }
    

    public get error(): string {
        return this._errorMessage;
    }

    public get result(): number {
        return this._result;
    }

    isNumber(token: TokenType): boolean {
        return !isNaN(Number(token));
    }

    isCellReference(token: TokenType): boolean {
        return Cell.isValidCellLabel(token);
    }

    getCellValue(token: TokenType): [number, string] {
        let cell = this._sheetMemory.getCellByLabel(token);
        let formula = cell.getFormula();
        let error = cell.getError();

        if (error !== "" && error !== ErrorMessages.emptyFormula) {
            return [0, error];
        }

        if (formula.length === 0) {
            return [0, ErrorMessages.invalidCell];
        }

        let value = cell.getValue();
        return [value, ""];
    }
}

export default FormulaEvaluator;