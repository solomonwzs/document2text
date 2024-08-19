.SUFFIXES: .cpp .c

NAME	= document2text

CXX			= g++
CXXFLAGS	= -Wall -fpic -g -c -std=c++20 -O2 \
			  -I. -I/usr/include/poppler \
			  -Wno-sign-compare -Wno-address-of-packed-member
CXXSRC		= $(wildcard ./*.cpp ./*/*.cpp ./*/*/*.cpp ./*/*/*/*.cpp ./*/*/*/*/*.cpp)
CXXOBJ		= $(CXXSRC:%.cpp=%-cpp.o)
CXXDEP		= $(CXXOBJ:%-cpp.o=%-cpp.d)

C		= gcc
CFLAGS	= -Wall -fpic -g -c
CSRC	= $(wildcard ./*.c)
COBJ	= $(CSRC:%.c=%-c.o)
CDEP	= $(COBJ:%-c.o=%-c.d)

LIBS	= -lpoppler -lzip

AR 		= ar
ARFLAGS	= rv
RANLIB	= ranlib

VALGRIND = valgrind

$(NAME).out: $(CXXOBJ) $(COBJ)
	@echo -e "\033[0;33m>>>\033[0m $@"
	@$(CXX) $(CXXOBJ) $(COBJ) $(LIBS) -o $@

lib$(NAME).so: $(CXXOBJ) $(COBJ)
	@echo -e "\033[0;33m>>>\033[0m $@"
	@$(CXX) -shared -Wl,-soname,lib$(NAME).so $(CXXOBJ) $(COBJ) -o $@

lib$(NAME).a: $(CXXOBJ) $(COBJ)
	@echo -e "\033[0;33m>>>\033[0m $@"
	@$(AR) $(ARFLAGS) $@ $(CXXOBJ) $(COBJ)
	@$(RANLIB) $@

-include $(CXXDEP)
-include $(CDEP)

.SECONDARY:
%-cpp.o: %.cpp
	@echo -e "\033[0;33m*\033[0m $< -> $@"
	@$(CXX) $(CXXFLAGS) $< -MMD -o $@

.SECONDARY:
%-c.o: %.c
	@echo -e "\033[0;33m*\033[0m $< -> $@"
	@$(C) $(CFLAGS) $< -MMD -o $@

.PHONY:
clean:
	-rm *.d *.o ./*/*.d ./*/*.o $(NAME).out lib$(NAME).so lib$(NAME).a

.PHONY:
mem_test: a.out
	@$(VALGRIND)	\
		--tool=memcheck \
		--leak-check=yes \
		--show-reachable=yes \
		--num-callers=20 \
		--track-fds=yes \
		./a.out
