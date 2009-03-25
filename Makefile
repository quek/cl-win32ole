clean:
	find ./ -name '*.fasl' -o -name '*~' -o -name '*.fas' -o -name '*.lib'| xargs rm
