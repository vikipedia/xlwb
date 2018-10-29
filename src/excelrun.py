from xlwb.xlspy import excelexec as ex



if __name__ == "__main__":
    args = ex.parse_args()
    ex.main(args.exceldata, args.inputs)
