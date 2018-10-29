from xlwb.xlspy import excelparser as ep


if __name__ == "__main__":
    args = ep.parse_args()
    ep.main(args.filename, args.output)
